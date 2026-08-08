import { OfficeKitCodecError } from "./office-kit-error.mjs";

export function assertTrustedImportedState(state, family) {
  if (!state) return;
  const sourceHash = String(state.source?.packageSha256 || "").toLowerCase();
  const snapshot = state.opaqueOpc?.sourcePackage;
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  if (!sourceHash || !snapshotHash || sourceHash !== snapshotHash || !snapshot?.data?.length) {
    throw new OfficeKitCodecError(`${family} source-bound export requires its validated source package snapshot.`, [], { code: "missing_source_package" });
  }
}
