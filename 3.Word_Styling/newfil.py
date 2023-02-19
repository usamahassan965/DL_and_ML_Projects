from typing import Dict
from lxml import etree as ET
import zipfile
import os
import shutil

def get_most_often_font_and_color(file_path):
    # Open the .docx file as a ZIP archive
    with zipfile.ZipFile(file_path, "r") as archive:
        # Read the contents of the styles.xml file
        styles_xml = archive.read("word/document.xml")

        # Parse the styles.xml contents as an XML document
        root = ET.fromstring(styles_xml)

        # Dictionary to store the most often used font-type and font-color for each style id
        style_dict = {}


        # Iterate over all the w:style elements in the document
        for p in root.findall(".//w:p", namespaces=root.nsmap):
            pPr = p.find(".//w:pPr", namespaces=root.nsmap)
            style = pPr.find(".//w:pStyle", namespaces=root.nsmap)
            if style != None:
                style_value = style.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"]

            rPr = pPr.find(".//w:rPr", namespaces=root.nsmap)
            Fonts = rPr.find(".//w:rFonts", namespaces=root.nsmap)
            color = rPr.find(".//w:color", namespaces=root.nsmap)
            if Fonts != None:
                font_ascii = Fonts.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii"]
            if color != None:
                color_value = color.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"]
            style_dict[style_value] = [font_ascii,color_value] 
            if  len(style_dict)== 1:
                break
            break
        print(style_value)
        print(font_ascii)
        print(color_value)
        print(style_dict)

get_most_often_font_and_color(r"C:\Users\HP\Desktop\Fiverr\73.Word_Styling\Source File A.docx")


def apply_most_font_and_color_to_styles(source_file_path, target_file_path):
    # Get most often font and color dictionary from source file
    most_often_font, most_often_color = 

    # Open the .docx file as a ZIP archive
    with zipfile.ZipFile(target_file_path, "a") as archive:
        # Read the contents of the styles.xml file
        styles_xml = archive.read("word/document.xml")

        # Parse the styles.xml contents as an XML document
        root = ET.fromstring(styles_xml)

        # Iterate over all the w:style elements in the document
        for p in root.findall(".//w:p", namespaces=root.nsmap):
            pPr = p.find(".//w:pPr", namespaces=root.nsmap)
            style = pPr.find(".//w:pStyle", namespaces=root.nsmap)
            if style != None:
                style_value = style.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"]
            if style_value in style_dict.items()[0]:
                    font_type_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
                    font_type_element.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii"] = font_type

            rPr = pPr.find(".//w:rPr", namespaces=root.nsmap)
            Fonts = rPr.find(".//w:rFonts", namespaces=root.nsmap)
            color = rPr.find(".//w:color", namespaces=root.nsmap)
            if Fonts != None:
                font_ascii = Fonts.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii"]
            if color != None:
                color_value = color.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"]

        for style_element in root.findall(".//w:style", namespaces=root.nsmap):
            # Extract the style ID and type
            style_id = style_element.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"]
            style_type = style_element.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type"]

            # Get the most often used font-type and font-color for the current style id
            if style_id in most_often_font:
                font_type = most_often_font[style_id]

                # Add/Update the w:rFonts element for the style
                rpr_element = style_element.find(".//w:rPr", namespaces=root.nsmap)
                if rpr_element is None:
                    rpr_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
                    style_element.append(rpr_element)

                font_type_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
                font_type_element.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii"] = font_type
                rpr_element.append(font_type_element)

            if style_id in most_often_color:
                font_color = most_often_color[style_id]

                # Add/Update the w:color element for the style
                rpr_element = style_element.find(".//w:rPr", namespaces=root.nsmap)
                if rpr_element is None:
                    rpr_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
                    style_element.append(rpr_element)

                font_color_element = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
                font_color_element.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"] = font_color
                rpr_element.append(font_color_element)

        # Write the modified XML back to the ZIP archive
        with archive.open("word/styles.xml", "w") as f:
            f.write(ET.tostring(root))

    # Create a backup of the target file with the suffix "_backup"
    backup_file_path = os.path.splitext(target_file_path)[0] + "_backup.docx"
    shutil.copy2(target_file_path, backup_file_path)

    print("Updated styles.xml written to", target_file_path)
    print("Backup created at", backup_file_path)

apply_most_font_and_color_to_styles(r'C:\Users\HP\Desktop\Fiverr\73.Word_Styling\Source File A.docx',r'C:\Users\HP\Desktop\Fiverr\73.Word_Styling\Target File A.docx')
