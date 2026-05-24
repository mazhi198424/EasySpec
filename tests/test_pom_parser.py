# tests/test_pom_parser.py
import tempfile
import os
from src.parsers.pom_parser import parse_pom


def test_parse_pom_returns_system_info(sample_pom_content):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
        f.write(sample_pom_content)
        tmp_path = f.name

    try:
        info = parse_pom(tmp_path)
        assert info.name == "VisionDemo"
        assert info.version == "1.0.0"
        assert info.packaging == "war"
    finally:
        os.unlink(tmp_path)


def test_parse_pom_no_name_fallback_to_artifact_id(sample_pom_content):
    content_no_name = sample_pom_content.replace('<name>VisionDemo</name>', '')
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
        f.write(content_no_name)
        tmp_path = f.name

    try:
        info = parse_pom(tmp_path)
        assert info.name == "vision-demo"
    finally:
        os.unlink(tmp_path)
