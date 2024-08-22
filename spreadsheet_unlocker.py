import os
import zipfile
import shutil
from lxml import etree

class ExcelWorkbookModifier:
    def __init__(self, file_path):
        self.original_xlsm_file_path = file_path
        self.cwd = os.path.dirname(file_path)
        self.name = os.path.basename(file_path)
        self.extract_to_path = os.path.join(self.cwd, f"{self.name.rsplit('.', 1)[0]}/extracted")
        self.worksheets_directory = os.path.join(self.extract_to_path, 'xl', 'worksheets')
        self.workbook_file_path = os.path.join(self.extract_to_path, 'xl', 'workbook.xml')
        self.zip_file_path = os.path.join(self.cwd, f"unlocked_{self.name.rsplit('.', 1)[0]}.zip")
        self.new_xlsm_file_path = os.path.join(self.cwd, f"unlocked_{self.name}")
        self.namespaces = {'default': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    def extract_file(self):
        os.makedirs(self.extract_to_path, exist_ok=True)
        with zipfile.ZipFile(self.original_xlsm_file_path, 'r') as zip_ref:
            zip_ref.extractall(self.extract_to_path)
        print("File extracted successfully!")

    def remove_sheet_protection(self):
        for i in range(1, 1000):
            filename = f'sheet{i}.xml'
            filepath = os.path.join(self.worksheets_directory, filename)
            if os.path.isfile(filepath):
                try:
                    parser = etree.XMLParser(remove_blank_text=True)
                    tree = etree.parse(filepath, parser)
                    root = tree.getroot()
                    sheet_protection_tags = root.xpath('.//default:sheetProtection[@algorithmName="SHA-512"]', namespaces=self.namespaces)
                    for tag in sheet_protection_tags:
                        parent = tag.getparent()
                        if parent is not None:
                            parent.remove(tag)
                    tree.write(filepath, pretty_print=True, xml_declaration=True, encoding='utf-8')
                except etree.XMLSyntaxError as e:
                    print(f"Reached end at {filename}: {e}")

    def remove_workbook_protection(self):
        if os.path.isfile(self.workbook_file_path):
            try:
                parser = etree.XMLParser(remove_blank_text=True)
                tree = etree.parse(self.workbook_file_path, parser)
                root = tree.getroot()
                workbook_protection_tag = root.find('.//default:workbookProtection', self.namespaces)
                if workbook_protection_tag is not None:
                    parent = workbook_protection_tag.getparent()
                    if parent is not None:
                        parent.remove(workbook_protection_tag)
                tree.write(self.workbook_file_path, pretty_print=True, xml_declaration=True, encoding='utf-8')
                print("Workbook protection removed successfully!")
            except etree.XMLSyntaxError as e:
                print(f"Failed to parse {self.workbook_file_path}: {e}")

    def zip_directory(self):
        def zipdir(path, ziph):
            for root, dirs, files in os.walk(path):
                for file in files:
                    filepath = os.path.join(root, file)
                    ziph.write(filepath, os.path.relpath(filepath, path))

        with zipfile.ZipFile(self.zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipdir(self.extract_to_path, zipf)
        print("Zip file created successfully!")

    def cleanup(self):
        shutil.rmtree(os.path.dirname(self.extract_to_path))
        print("Removed trace files successfully!")
        os.rename(self.zip_file_path, self.new_xlsm_file_path)
        print("File renamed successfully!")

    def modify_workbook(self):
        self.extract_file()
        self.remove_sheet_protection()
        self.remove_workbook_protection()
        self.zip_directory()
        self.cleanup()
