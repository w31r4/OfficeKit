#!/bin/sh

set -eu
umask 077

OFFICE_KIT_VERSION=1.1.0
OFFICE_KIT_REPOSITORY=w31r4/OfficeKit

fail() {
  printf '%s\n' "OfficeKit installer: $*" >&2
  exit 1
}

command_exists() {
  command -v "$1" >/dev/null 2>&1
}

detect_target() {
  kernel=$(uname -s)
  machine=$(uname -m)
  case "$kernel-$machine" in
    Darwin-arm64) printf '%s\n' darwin-arm64 ;;
    Linux-x86_64|Linux-amd64) printf '%s\n' linux-x64 ;;
    *) fail "no self-contained build is available for $kernel/$machine." ;;
  esac
}

file_size() {
  wc -c < "$1" | tr -d '[:space:]'
}

file_sha256() {
  if command_exists sha256sum; then
    sha256sum "$1" | awk '{print $1}'
  elif command_exists shasum; then
    shasum -a 256 "$1" | awk '{print $1}'
  else
    fail "sha256sum or shasum is required."
  fi
}

download() {
  url=$1
  output=$2
  if command_exists curl; then
    curl --fail --location --silent --show-error "$url" --output "$output"
  elif command_exists wget; then
    wget --quiet "$url" --output-document="$output"
  else
    fail "curl or wget is required."
  fi
}

configure_path() {
  if [ "${OFFICE_KIT_INSTALL_TEST:-0}" = "1" ] &&
    [ "${OFFICE_KIT_TEST_CONFIGURE_PATH:-0}" != "1" ]; then
    return 0
  fi
  if [ "$bin_root" != "$HOME/.local/bin" ]; then
    return 0
  fi
  case "$bin_root" in
    *'"'*|*'`'*) fail "installation path contains unsupported shell characters." ;;
  esac
  shell_name=${SHELL##*/}
  profile=
  case "$shell_name" in
    zsh) profile="$HOME/.zshrc" ;;
    bash)
      if [ -f "$HOME/.bash_profile" ]; then
        profile="$HOME/.bash_profile"
      else
        profile="$HOME/.bashrc"
      fi
      ;;
    *) return 0 ;;
  esac
  path_line="export PATH=\"$bin_root:\$PATH\""
  if [ ! -e "$profile" ]; then
    (umask 077 && : > "$profile") || fail "could not create $profile."
  fi
  [ -f "$profile" ] && [ ! -L "$profile" ] ||
    fail "$profile must be a regular non-symlink shell profile."
  if ! grep -F "$path_line" "$profile" >/dev/null 2>&1; then
    printf '\n# OfficeKit\n%s\n' "$path_line" >> "$profile" ||
      fail "could not update $profile."
  fi
}

validate_listing() {
  listing=$1
  expected_root=$2
  entry_count=0
  while IFS= read -r raw_entry; do
    [ -n "$raw_entry" ] || continue
    entry_count=$((entry_count + 1))
    entry=${raw_entry%/}
    case "$entry" in
      /*|../*|*/../*|*/..|.|..) fail "archive contains an unsafe path: $raw_entry" ;;
    esac
    case "$entry" in
      "$expected_root"|"$expected_root"/*) ;;
      *) fail "archive entry is outside $expected_root: $raw_entry" ;;
    esac
  done < "$listing"
  [ "$entry_count" -gt 0 ] || fail "archive contains no entries."
}

validate_installation() {
  probe_root=$1
  receipt_mode=${2:-}
  [ -f "$probe_root/standalone-manifest.json" ] || fail "standalone manifest is missing."
  [ ! -L "$probe_root/standalone-manifest.json" ] || fail "standalone manifest cannot be a symlink."
  [ -x "$probe_root/runtime/node/bin/node" ] || fail "bundled Node executable is missing."
  [ ! -L "$probe_root/runtime/node/bin/node" ] || fail "bundled Node executable cannot be a symlink."
  [ -x "$probe_root/bin/officekit" ] || fail "OfficeKit launcher is missing."
  [ ! -L "$probe_root/bin/officekit" ] || fail "OfficeKit launcher cannot be a symlink."
  [ -f "$probe_root/lib/verify-install.mjs" ] || fail "installation verifier is missing."
  [ ! -L "$probe_root/lib/verify-install.mjs" ] || fail "installation verifier cannot be a symlink."
  if [ "$receipt_mode" = "allow-receipt" ]; then
    "$probe_root/runtime/node/bin/node" \
      "$probe_root/lib/verify-install.mjs" \
      "$probe_root" "$OFFICE_KIT_VERSION" "$target" --allow-receipt \
      >/dev/null ||
      fail "installed file integrity verification failed."
  else
    "$probe_root/runtime/node/bin/node" \
      "$probe_root/lib/verify-install.mjs" \
      "$probe_root" "$OFFICE_KIT_VERSION" "$target" \
      >/dev/null ||
      fail "archive file integrity verification failed."
  fi
  actual_version=$(
    "$probe_root/runtime/node/bin/node" \
      "$probe_root/app/node_modules/office-kit/bin/officekit.mjs" \
      --version
  ) || fail "bundled OfficeKit failed its version probe."
  [ "$actual_version" = "$OFFICE_KIT_VERSION" ] ||
    fail "bundle version is $actual_version; expected $OFFICE_KIT_VERSION."
}

if [ "${OFFICE_KIT_INSTALL_TEST:-0}" = "1" ]; then
  target=${OFFICE_KIT_TEST_TARGET:-$(detect_target)}
else
  target=$(detect_target)
fi
case "$target" in
  darwin-arm64)
    expected_sha256=bec099f3e8c4e5b98a988c36d0054dc2e890002d35a62c9ddfbe82039c8cf0ff
    expected_size=89097571
    ;;
  linux-x64)
    expected_sha256=71dc7e6be2b58991953b1d863cfb42ee1f2fc3a1d55731292ae1c1bf67e9897d
    expected_size=94223216
    ;;
  *) fail "unsupported target $target." ;;
esac

asset="office-kit-$OFFICE_KIT_VERSION-$target.tar.gz"
archive_url="https://github.com/$OFFICE_KIT_REPOSITORY/releases/download/v$OFFICE_KIT_VERSION/$asset"

if [ "${OFFICE_KIT_INSTALL_TEST:-0}" = "1" ]; then
  [ -n "${OFFICE_KIT_TEST_ARCHIVE:-}" ] || fail "OFFICE_KIT_TEST_ARCHIVE is required in test mode."
  [ -n "${OFFICE_KIT_TEST_SHA256:-}" ] || fail "OFFICE_KIT_TEST_SHA256 is required in test mode."
  [ -n "${OFFICE_KIT_TEST_SIZE:-}" ] || fail "OFFICE_KIT_TEST_SIZE is required in test mode."
  expected_sha256=$OFFICE_KIT_TEST_SHA256
  expected_size=$OFFICE_KIT_TEST_SIZE
fi

case "$expected_sha256" in
  *[!0-9a-f]*|'') fail "release SHA-256 is not finalized." ;;
esac
[ "${#expected_sha256}" -eq 64 ] || fail "release SHA-256 is not finalized."
case "$expected_size" in
  *[!0-9]*|'') fail "release size is not finalized." ;;
esac

install_root=${OFFICE_KIT_HOME:-"$HOME/.office-kit"}
bin_root=${OFFICE_KIT_BIN_DIR:-"$HOME/.local/bin"}
versions_root="$install_root/versions"
version_root="$versions_root/$OFFICE_KIT_VERSION"
current_link="$install_root/current"
command_link="$bin_root/officekit"

mkdir -p "$versions_root" "$bin_root"
temporary=$(mktemp -d "$install_root/.install.XXXXXX") ||
  fail "could not create an installation transaction."
cleanup() {
  rm -rf "$temporary"
}
trap cleanup EXIT HUP INT TERM

archive="$temporary/$asset"
if [ "${OFFICE_KIT_INSTALL_TEST:-0}" = "1" ]; then
  cp "$OFFICE_KIT_TEST_ARCHIVE" "$archive"
else
  download "$archive_url" "$archive"
fi

actual_size=$(file_size "$archive")
[ "$actual_size" = "$expected_size" ] ||
  fail "archive size is $actual_size; expected $expected_size."
actual_sha256=$(file_sha256 "$archive")
[ "$actual_sha256" = "$expected_sha256" ] ||
  fail "archive SHA-256 is $actual_sha256; expected $expected_sha256."

archive_root="office-kit-$OFFICE_KIT_VERSION-$target"
listing="$temporary/archive.list"
tar -tzf "$archive" > "$listing" || fail "archive listing failed."
validate_listing "$listing" "$archive_root"

if [ -e "$version_root" ] || [ -L "$version_root" ]; then
  [ -d "$version_root" ] && [ ! -L "$version_root" ] ||
    fail "existing version path is not a real directory: $version_root."
  receipt="$version_root/.office-kit-install-receipt"
  [ -f "$receipt" ] && [ ! -L "$receipt" ] ||
    fail "existing version is missing its installation receipt."
  expected_receipt="office-kit.standalone-install.v1 $OFFICE_KIT_VERSION $target $expected_sha256 $expected_size"
  [ "$(cat "$receipt")" = "$expected_receipt" ] ||
    fail "existing version receipt does not match this release."
  validate_installation "$version_root" allow-receipt
else
  extraction="$temporary/extracted"
  mkdir "$extraction"
  tar -xzf "$archive" -C "$extraction" || fail "archive extraction failed."
  candidate="$extraction/$archive_root"
  [ -d "$candidate" ] && [ ! -L "$candidate" ] ||
    fail "archive root is not a real directory."
  validate_installation "$candidate"
  printf '%s\n' \
    "office-kit.standalone-install.v1 $OFFICE_KIT_VERSION $target $expected_sha256 $expected_size" \
    > "$candidate/.office-kit-install-receipt"
  mv "$candidate" "$version_root" ||
    fail "could not publish version $OFFICE_KIT_VERSION."
fi

if [ -e "$command_link" ] && [ ! -L "$command_link" ]; then
  fail "$command_link already exists and is not an OfficeKit symlink."
fi
if [ -L "$command_link" ]; then
  existing_command_target=$(readlink "$command_link")
  [ "$existing_command_target" = "$current_link/bin/officekit" ] ||
    fail "$command_link points to another installation."
fi
if [ -e "$current_link" ] && [ ! -L "$current_link" ]; then
  fail "$current_link already exists and is not an OfficeKit symlink."
fi
if [ -L "$current_link" ]; then
  existing_current_target=$(readlink "$current_link")
  case "$existing_current_target" in
    versions/*)
      current_version=${existing_current_target#versions/}
      case "$current_version" in
        *[!A-Za-z0-9._-]*|''|.|..) fail "$current_link has an invalid managed version target." ;;
      esac
      ;;
    *) fail "$current_link points outside the managed versions directory." ;;
  esac
fi

command_next="$bin_root/.officekit.next.$$"
current_next="$install_root/.current.next.$$"
rm -f "$command_next" "$current_next"
ln -s "$current_link/bin/officekit" "$command_next" ||
  fail "could not prepare the OfficeKit command."
mv -f "$command_next" "$command_link" ||
  fail "could not publish the OfficeKit command."
ln -s "versions/$OFFICE_KIT_VERSION" "$current_next" ||
  fail "could not prepare the active OfficeKit version."
mv -f "$current_next" "$current_link" ||
  fail "could not activate OfficeKit $OFFICE_KIT_VERSION."

configure_path

printf '%s\n' "OfficeKit $OFFICE_KIT_VERSION installed at $version_root"
printf '%s\n' "Command: $command_link"
case ":$PATH:" in
  *":$bin_root:"*) ;;
  *) printf '%s\n' "Open a new terminal, then run officekit." ;;
esac
