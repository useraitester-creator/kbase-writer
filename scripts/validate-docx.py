"""Check the DOCX package, embedded resources and relationships; not pagination."""
import argparse
import json
import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import unquote
from zipfile import ZipFile


def validate(path, expected=None):
    with ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise ValueError("Повреждён ZIP-контейнер DOCX")
        names = set(archive.namelist())
        for required in ("[Content_Types].xml", "word/document.xml", "_rels/.rels"):
            if required not in names:
                raise ValueError(f"Нет обязательной части: {required}")
        xml = {name: ET.fromstring(archive.read(name)) for name in names
               if name.endswith((".xml", ".rels"))}
        rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        for name, root in xml.items():
            if not name.endswith(".rels"):
                continue
            base = posixpath.dirname(posixpath.dirname(name))
            for relation in root.findall(f"{rel_ns}Relationship"):
                if relation.get("TargetMode") == "External":
                    continue
                target = unquote(relation.get("Target", "").split("#", 1)[0])
                resolved = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
                if resolved not in names:
                    raise ValueError(f"Отсутствует ресурс: {resolved}")
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
              "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        root = xml["word/document.xml"]
        text = " ".join(node.text or "" for node in root.findall(".//w:t", ns))
        if not text.strip():
            raise ValueError("В документе нет текста")
        figures = len(root.findall(".//a:blip", ns))
        if expected is not None and figures != expected:
            raise ValueError(f"Ожидалось {expected} изображений, найдено {figures}")
        return {"valid": True, "bytes": Path(path).stat().st_size,
                "words": len(text.split()), "figures": figures,
                "tables": len(root.findall(".//w:tbl", ns)),
                "paginationChecked": False}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file")
    parser.add_argument("--expect-images", type=int)
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.file, args.expect_images), ensure_ascii=False))
    except Exception as error:
        parser.exit(1, f"Ошибка DOCX: {error}\n")
