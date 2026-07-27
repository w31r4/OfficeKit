/*
 * Minimal Windows launcher for the capability-pack-owned Ghostscript console
 * binary.  Ghostscript needs an explicit resource search path after its
 * executable is relocated, so this launcher derives every path from the pack
 * root, never invokes a shell, and removes ambient GS configuration before
 * starting the real console program.
 */

#define UNICODE
#define _UNICODE

#include <windows.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

#define OFFICEKIT_MAX_PATH 32768

typedef struct {
  wchar_t *value;
  size_t length;
  size_t capacity;
} wide_buffer;

static int reserve(wide_buffer *buffer, size_t additional) {
  if (additional > SIZE_MAX - buffer->length - 1) return 0;
  size_t required = buffer->length + additional + 1;
  if (required <= buffer->capacity) return 1;
  size_t next = buffer->capacity ? buffer->capacity : 256;
  while (next < required) {
    if (next > SIZE_MAX / 2) {
      next = required;
      break;
    }
    next *= 2;
  }
  wchar_t *replacement = (wchar_t *)realloc(buffer->value, next * sizeof(wchar_t));
  if (!replacement) return 0;
  buffer->value = replacement;
  buffer->capacity = next;
  return 1;
}

static int append_n(wide_buffer *buffer, const wchar_t *text, size_t length) {
  if (!reserve(buffer, length)) return 0;
  memcpy(buffer->value + buffer->length, text, length * sizeof(wchar_t));
  buffer->length += length;
  buffer->value[buffer->length] = L'\0';
  return 1;
}

static int append_char(wide_buffer *buffer, wchar_t character) {
  return append_n(buffer, &character, 1);
}

static int append_repeat(wide_buffer *buffer, wchar_t character, size_t count) {
  if (!reserve(buffer, count)) return 0;
  for (size_t index = 0; index < count; index += 1) buffer->value[buffer->length + index] = character;
  buffer->length += count;
  buffer->value[buffer->length] = L'\0';
  return 1;
}

/* Quote an argv element according to CreateProcess/CommandLineToArgvW rules. */
static int append_quoted(wide_buffer *buffer, const wchar_t *argument) {
  if (!append_char(buffer, L'\"')) return 0;
  size_t slashes = 0;
  for (const wchar_t *cursor = argument; *cursor; cursor += 1) {
    if (*cursor == L'\\') {
      slashes += 1;
      continue;
    }
    if (*cursor == L'\"') {
      if (!append_repeat(buffer, L'\\', slashes * 2 + 1) || !append_char(buffer, L'\"')) return 0;
      slashes = 0;
      continue;
    }
    if (!append_repeat(buffer, L'\\', slashes) || !append_char(buffer, *cursor)) return 0;
    slashes = 0;
  }
  return append_repeat(buffer, L'\\', slashes * 2) && append_char(buffer, L'\"');
}

static wchar_t *copy_path(const wchar_t *source) {
  size_t length = wcslen(source);
  wchar_t *result = (wchar_t *)calloc(length + 1, sizeof(wchar_t));
  if (result) memcpy(result, source, (length + 1) * sizeof(wchar_t));
  return result;
}

static int parent_directory(wchar_t *path) {
  wchar_t *separator = wcsrchr(path, L'\\');
  wchar_t *slash = wcsrchr(path, L'/');
  if (slash && (!separator || slash > separator)) separator = slash;
  if (!separator || separator == path) return 0;
  *separator = L'\0';
  return 1;
}

static wchar_t *join_path(const wchar_t *root, const wchar_t *relative) {
  size_t root_length = wcslen(root);
  size_t relative_length = wcslen(relative);
  if (root_length > SIZE_MAX - relative_length - 2) return NULL;
  wchar_t *result = (wchar_t *)calloc(root_length + relative_length + 2, sizeof(wchar_t));
  if (!result) return NULL;
  memcpy(result, root, root_length * sizeof(wchar_t));
  result[root_length] = L'\\';
  memcpy(result + root_length + 1, relative, (relative_length + 1) * sizeof(wchar_t));
  return result;
}

static wchar_t *join_list(const wchar_t *first, const wchar_t *second) {
  size_t first_length = wcslen(first);
  size_t second_length = wcslen(second);
  if (first_length > SIZE_MAX - second_length - 2) return NULL;
  wchar_t *result = (wchar_t *)calloc(first_length + second_length + 2, sizeof(wchar_t));
  if (!result) return NULL;
  memcpy(result, first, first_length * sizeof(wchar_t));
  result[first_length] = L';';
  memcpy(result + first_length + 1, second, (second_length + 1) * sizeof(wchar_t));
  return result;
}

static int is_regular_non_reparse_file(const wchar_t *path) {
  DWORD attributes = GetFileAttributesW(path);
  return attributes != INVALID_FILE_ATTRIBUTES
    && !(attributes & FILE_ATTRIBUTE_DIRECTORY)
    && !(attributes & FILE_ATTRIBUTE_REPARSE_POINT);
}

static void report_failure(const wchar_t *message) {
  fwprintf(stderr, L"office-kit Ghostscript launcher: %ls (Windows error %lu)\n", message, GetLastError());
}

int wmain(int argc, wchar_t **argv) {
  wchar_t executable[OFFICEKIT_MAX_PATH];
  DWORD executable_length = GetModuleFileNameW(NULL, executable, OFFICEKIT_MAX_PATH);
  if (executable_length == 0 || executable_length >= OFFICEKIT_MAX_PATH) {
    report_failure(L"could not resolve launcher path");
    return 2;
  }

  wchar_t *root = copy_path(executable);
  if (!root || !parent_directory(root) || !parent_directory(root)) {
    report_failure(L"launcher is not below a capability-pack bin directory");
    free(root);
    return 2;
  }
  wchar_t *target = join_path(root, L"bin\\gswin64c.exe");
  wchar_t *bin = join_path(root, L"bin");
  wchar_t *dlls = join_path(root, L"DLLs");
  wchar_t *resource = join_path(root, L"share\\ghostscript\\Resource");
  wchar_t *library = join_path(root, L"share\\ghostscript\\lib");
  const wchar_t *system_root_value = _wgetenv(L"SystemRoot");
  wchar_t *system_root = copy_path(system_root_value && *system_root_value ? system_root_value : L"C:\\Windows");
  if (!target || !bin || !dlls || !resource || !library || !system_root || !is_regular_non_reparse_file(target)) {
    report_failure(L"capability-pack Ghostscript runtime is missing or unsafe");
    free(root);
    free(target);
    free(bin);
    free(dlls);
    free(resource);
    free(library);
    free(system_root);
    return 2;
  }

  wide_buffer command = { 0 };
  int assembled = append_quoted(&command, target);
  for (int index = 1; assembled && index < argc; index += 1) {
    assembled = append_char(&command, L' ') && append_quoted(&command, argv[index]);
  }
  wchar_t *system32 = join_path(system_root, L"System32");
  wchar_t *gs_lib = join_list(resource, library);
  size_t path_length = wcslen(bin) + wcslen(root) + wcslen(dlls) + (system32 ? wcslen(system32) : 0) + 4;
  wchar_t *runtime_path = system32 ? (wchar_t *)calloc(path_length, sizeof(wchar_t)) : NULL;
  if (!assembled || !system32 || !gs_lib || !runtime_path) {
    report_failure(L"could not allocate Ghostscript command or environment");
    free(command.value);
    free(root);
    free(target);
    free(bin);
    free(dlls);
    free(resource);
    free(library);
    free(system_root);
    free(system32);
    free(gs_lib);
    free(runtime_path);
    return 2;
  }
  swprintf(runtime_path, path_length, L"%ls;%ls;%ls;%ls", bin, root, dlls, system32);

  /* Do not allow a caller's Ghostscript configuration to select another
   * resource tree, device, or startup option. */
  SetEnvironmentVariableW(L"GS_OPTIONS", NULL);
  SetEnvironmentVariableW(L"GS_DEVICE", NULL);
  SetEnvironmentVariableW(L"GS_LIB", gs_lib);
  SetEnvironmentVariableW(L"GS_FONTPATH", NULL);
  SetEnvironmentVariableW(L"SystemRoot", system_root);
  SetEnvironmentVariableW(L"PATH", runtime_path);

  STARTUPINFOW startup = { 0 };
  PROCESS_INFORMATION process = { 0 };
  startup.cb = sizeof(startup);
  if (!CreateProcessW(target, command.value, NULL, NULL, FALSE, 0, NULL, root, &startup, &process)) {
    report_failure(L"could not start capability-pack Ghostscript runtime");
    free(command.value);
    free(root);
    free(target);
    free(bin);
    free(dlls);
    free(resource);
    free(library);
    free(system_root);
    free(system32);
    free(gs_lib);
    free(runtime_path);
    return 2;
  }
  WaitForSingleObject(process.hProcess, INFINITE);
  DWORD exit_code = 2;
  if (!GetExitCodeProcess(process.hProcess, &exit_code)) report_failure(L"could not read Ghostscript process exit status");
  CloseHandle(process.hThread);
  CloseHandle(process.hProcess);
  free(command.value);
  free(root);
  free(target);
  free(bin);
  free(dlls);
  free(resource);
  free(library);
  free(system_root);
  free(system32);
  free(gs_lib);
  free(runtime_path);
  return (int)exit_code;
}
