// Limits that keep a hostile workbook from exhausting the editor. They live apart from the reader
// so a caller can refuse a file without pulling the deflate codec into its bundle.
export const OFFICE_ZIP_MAX_ENTRIES = 4000
export const OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
