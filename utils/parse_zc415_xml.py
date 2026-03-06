from xml.dom import minidom
from typing import Dict, Optional, IO


def parse_zc415_xml(file_stream: IO) -> Dict[str, str]:
    """
    解析 ZC415 XML 文件流，提取 (lrn → ucr) 映射

    :param file_stream: 打开的 XML 文件流对象
    :return: dict {lrn: ucr}
    """
    mapping = {}
    try:
        doc = minidom.parse(file_stream)
        declarations = doc.getElementsByTagName("Declaration")

        for decl in declarations:
            lrn = get_text_content_by_tag(decl, "lrn")
            goods_shipments = decl.getElementsByTagName("GoodsShipment")
            if goods_shipments:
                ucr = get_text_content_by_tag(goods_shipments[0], "referenceNumberUCR")
                if lrn and ucr:
                    mapping[lrn] = ucr
    except Exception as e:
        print("Error parsing ZC415 XML:", e)

    return mapping


def get_text_content_by_tag(parent, tag_name: str) -> Optional[str]:
    """获取指定节点下第一个 <tag_name> 的文本内容"""
    elements = parent.getElementsByTagName(tag_name)
    if elements and elements[0].firstChild:
        return elements[0].firstChild.nodeValue.strip()
    return None
