import json
from xml.dom import minidom
from typing import List, Tuple, Optional, IO


def parse_zc429_xml(file_stream: IO) -> List[Tuple[str, str]]:
    """
    解析 ZC429 XML 文件流，提取所有 <referenceNumberUCR> 与其对应的 <mrn>，返回元组列表。

    :param file_stream: 打开的 XML 文件流对象
    :return: List of (referenceNumberUCR, mrn)
    """
    release_info_list = []
    try:
        doc = minidom.parse(file_stream)
        print("Root element:", doc.documentElement.nodeName)
        ref_elements = doc.getElementsByTagName("referenceNumberUCR")
        print("Number of <referenceNumberUCR> nodes found:", len(ref_elements))

        for ref in ref_elements:
            reference_number_ucr = ref.firstChild.nodeValue.strip() if ref.firstChild else None
            mrn = find_mrn_upwards(ref.parentNode)

            print(f"ReferenceNumberUCR: {reference_number_ucr}, MRN: {mrn}")
            if reference_number_ucr and mrn:
                release_info_list.append((reference_number_ucr, mrn))

    except Exception as e:
        print("Error parsing XML:", e)

    return release_info_list

def parse_zc429_json(file_stream: IO) -> List[Tuple[str, str]]:
    """
    解析 ZC429 JSON 文件流，提取所有 referenceNumberUCR 与其对应的 mrn，
    返回元组列表，结构与 parse_zc429_xml 完全一致。

    :param file_stream: 打开的 JSON 文件流对象
    :return: List of (referenceNumberUCR, mrn)
    """
    release_info_list = []

    try:
        data = json.load(file_stream)

        hub = data.get("p:ZC429HUB", {})
        declarations = hub.get("Declaration")

        if isinstance(declarations, dict):
            declarations = [declarations]
        elif not isinstance(declarations, list):
            print(f"❌ Invalid Declaration type: {type(declarations)}")
            return release_info_list

        print("✅ Number of Declaration nodes found:", len(declarations))

        # ✅ 3️⃣ 统一循环提取
        for declaration in declarations:
            goods_shipment = declaration.get("GoodsShipment", {})

            reference_number_ucr = (
                goods_shipment.get("referenceNumberUCR", {})
                .get("value")
            )

            mrn = declaration.get("mrn", {}).get("value")

            if reference_number_ucr:
                reference_number_ucr = reference_number_ucr.strip()

            if mrn:
                mrn = mrn.strip()

            print(f"✅ ReferenceNumberUCR: {reference_number_ucr}, MRN: {mrn}")

            if reference_number_ucr and mrn:
                release_info_list.append((reference_number_ucr, mrn))

    except Exception as e:
        print("❌ Error parsing JSON:", e)

    return release_info_list

# if __name__ == '__main__':
#     with open("G:\\ZC429.xml", 'r', encoding='utf-8') as f:
#         x = parse_zc429_json(f)
#         print(x)

def find_mrn_upwards(node) -> Optional[str]:
    """
    向上查找包含 <mrn> 的父节点，并提取其文本内容。
    """
    while node:
        if node.nodeType == node.ELEMENT_NODE:
            mrn = get_text_content_by_tag(node, "mrn")
            if mrn:
                return mrn
        node = node.parentNode
    return None


def get_text_content_by_tag(parent, tag_name: str) -> Optional[str]:
    """
    获取指定节点下的第一个 <tag_name> 的文本内容。
    """
    elements = parent.getElementsByTagName(tag_name)
    if elements and elements[0].firstChild:
        return elements[0].firstChild.nodeValue.strip()
    return None

