import os
import re
import sys
import tempfile
import zipfile


def read_text(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def write_text(path, text):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def main():
    if len(sys.argv) != 4:
        raise SystemExit("usage: make-pptx.py <template.pptx> <out.pptx> <addin-id>")
    template, out, addin_id = sys.argv[1:]
    version = "1.0.0.1"
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(template, "r") as z:
            z.extractall(tmp)

        webext = os.path.join(tmp, "ppt", "webextensions", "webextension.xml")
        xml = read_text(webext)
        xml = re.sub(
            r'(<we:webextension\b[^>]*\sid=")\{?[0-9A-Fa-f-]{36}\}?"',
            lambda m: f'{m.group(1)}{{{addin_id}}}"',
            xml,
            count=1,
        )
        xml = re.sub(
            r'(<we:reference\b[^>]*\sid=")[0-9A-Fa-f-]{36}"',
            lambda m: f'{m.group(1)}{addin_id}"',
            xml,
            count=1,
        )
        xml = re.sub(
            r'(<we:reference\b[^>]*\sversion=")[^"]+"',
            lambda m: f'{m.group(1)}{version}"',
            xml,
            count=1,
        )
        write_text(webext, xml)

        # Make the task pane a little wider for Korean text.
        taskpanes = os.path.join(tmp, "ppt", "webextensions", "taskpanes.xml")
        tpxml = read_text(taskpanes).replace('width="350"', 'width="390"')
        write_text(taskpanes, tpxml)

        os.makedirs(os.path.dirname(out), exist_ok=True)
        if os.path.exists(out):
            os.remove(out)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp):
                for name in files:
                    full = os.path.join(root, name)
                    rel = os.path.relpath(full, tmp).replace(os.sep, "/")
                    z.write(full, rel)


if __name__ == "__main__":
    main()
