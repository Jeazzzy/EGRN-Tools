import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


class DataProcessor:
    @staticmethod
    def extract_data_from_zip(zip_path):
        try:
            with zipfile.ZipFile(zip_path) as archive:
                xml_files = [name for name in archive.namelist() if name.lower().endswith(".xml")]
                if not xml_files:
                    return None, None, "XML файл не найден"
                with archive.open(xml_files[0]) as source:
                    root = ET.parse(source).getroot()
                district = index = ""
                for element in root.iter():
                    tag = element.tag.split("}")[-1]
                    if tag == "cadastral_district":
                        district = element.text or ""
                    elif tag == "index":
                        index = element.text or ""
                if not district or not index:
                    return None, None, "Данные не найдены"
                return district, index, None
        except zipfile.BadZipFile:
            return None, None, "Неверный ZIP-файл"
        except ET.ParseError:
            return None, None, "Ошибка парсинга XML"
        except Exception as error:
            return None, None, str(error)

    @staticmethod
    def process_folder(base_path, progress_callback=None):
        base_path, results = Path(base_path), {}

        def process_settlement(path):
            data = {}
            for index_folder in path.iterdir():
                if not index_folder.is_dir():
                    continue
                archives = list(index_folder.glob("*.zip"))
                if not archives:
                    continue
                district, index, error = DataProcessor.extract_data_from_zip(archives[0])
                data[index_folder.name] = (
                    {"district": None, "error": error}
                    if error else {"district": district, "index": index}
                )
            return data

        direct = any(item.is_dir() and any(item.glob("*.zip")) for item in base_path.iterdir())
        if direct:
            data = process_settlement(base_path)
            if data:
                results[base_path.name] = data
            if progress_callback:
                progress_callback(1, 1)
        else:
            settlements = [path for path in base_path.iterdir() if path.is_dir()]
            for done, settlement in enumerate(settlements, 1):
                data = process_settlement(settlement)
                if data:
                    results[settlement.name] = data
                if progress_callback:
                    progress_callback(done, len(settlements))
        return results
