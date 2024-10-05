import mime from "mime-types";
import path from "path";

export const getMimeType = (filename: string): string | false => {
    const ext = path.extname(filename);
    return mime.lookup(ext);
};
