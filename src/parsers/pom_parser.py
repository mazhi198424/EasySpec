# src/parsers/pom_parser.py
import xml.etree.ElementTree as ET
from src.models import SystemInfo

NS = {'m': 'http://maven.apache.org/POM/4.0.0'}


def parse_pom(filepath: str) -> SystemInfo:
    tree = ET.parse(filepath)
    root = tree.getroot()

    name = _find_text(root, 'm:name', '')
    artifact_id = _find_text(root, 'm:artifactId', 'unknown')
    version = _find_text(root, 'm:version', '0.0.0')
    packaging = _find_text(root, 'm:packaging', 'jar')

    return SystemInfo(
        name=name or artifact_id,
        version=version,
        packaging=packaging
    )


def _find_text(root, tag, default):
    el = root.find(tag, NS)
    return el.text.strip() if el is not None and el.text else default
