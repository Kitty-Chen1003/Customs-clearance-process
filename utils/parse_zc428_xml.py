from xml.dom import minidom
from typing import Dict, Optional, IO


def parse_zc428_xml(file_stream: IO) -> Dict[str, str]:
    """
    解析 ZC428 XML 文件流，提取 (lrn → mrn) 映射
    """
    mapping = {}
    try:
        doc = minidom.parse(file_stream)
        declarations = doc.getElementsByTagName("Declaration")

        for decl in declarations:
            lrn = get_text_content_by_tag(decl, "lrn")
            mrn = get_text_content_by_tag(decl, "mrn")
            if lrn and mrn:
                mapping[lrn] = mrn
    except Exception as e:
        print("Error parsing ZC428 XML:", e)

    return mapping


def get_text_content_by_tag(parent, tag_name: str) -> Optional[str]:
    elements = parent.getElementsByTagName(tag_name)
    if elements and elements[0].firstChild:
        return elements[0].firstChild.nodeValue.strip()
    return None
