"""
Data Loaders Module - Format-agnostic data loading
Supports Excel, JSON, CSV, and other formats
"""

import pandas as pd
import json
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Union
import os
from pathlib import Path


class BaseDataLoader(ABC):
    """Abstract base class for data loaders"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path).split('.')[0]
        
    @abstractmethod
    def load_data(self) -> pd.DataFrame:
        """Load data and return as pandas DataFrame"""
        pass
    
    @abstractmethod
    def get_supported_extensions(self) -> List[str]:
        """Return list of supported file extensions"""
        pass


class ExcelDataLoader(BaseDataLoader):
    """Data loader for Excel files (.xlsx, .xls)"""
    
    def get_supported_extensions(self) -> List[str]:
        return ['.xlsx', '.xls']
    
    def load_data(self) -> pd.DataFrame:
        """Load Excel data with multiple fallback methods"""
        methods = [
            ("Default pandas", lambda: pd.read_excel(self.file_path)),
            ("All sheets", self._try_all_sheets),
            ("Named sheet", self._try_named_sheets),
            ("Openpyxl engine", lambda: pd.read_excel(self.file_path, engine='openpyxl')),
            ("Header None", lambda: pd.read_excel(self.file_path, header=None)),
        ]
        
        for method_name, method_func in methods:
            try:
                print(f"Trying {method_name}...")
                df = method_func()
                if df is not None and not df.empty:
                    print(f"✅ {method_name} successful! Shape: {df.shape}")
                    return df
            except Exception as e:
                print(f"❌ {method_name} failed: {e}")
                continue
        
        raise Exception(f"All Excel loading methods failed for {self.file_name}")
    
    def _try_all_sheets(self):
        """Try to read all sheets and return the first non-empty one"""
        all_sheets = pd.read_excel(self.file_path, sheet_name=None)
        for sheet_name, df in all_sheets.items():
            if not df.empty:
                print(f"Found data in sheet: '{sheet_name}'")
                return df
        return None
    
    def _try_named_sheets(self):
        """Try common sheet names"""
        common_names = [self.file_name, 'Summary', 'Data', 'Sheet1', 'Sheet 1', 'Main']
        for name in common_names:
            try:
                df = pd.read_excel(self.file_path, sheet_name=name)
                if not df.empty:
                    print(f"Found data in sheet: '{name}'")
                    return df
            except:
                continue
        return None


class JSONDataLoader(BaseDataLoader):
    """Data loader for JSON files (.json)"""
    
    def get_supported_extensions(self) -> List[str]:
        return ['.json']
    
    def load_data(self) -> pd.DataFrame:
        """Load JSON data with multiple format support"""
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Handle different JSON structures
            if isinstance(data, list):
                # Array of objects
                df = pd.DataFrame(data)
            elif isinstance(data, dict):
                if 'data' in data:
                    # Nested structure with 'data' key
                    df = pd.DataFrame(data['data'])
                elif 'records' in data:
                    # Nested structure with 'records' key
                    df = pd.DataFrame(data['records'])
                elif 'results' in data:
                    # Nested structure with 'results' key
                    df = pd.DataFrame(data['results'])
                else:
                    # Try to convert dict directly (single record or key-value pairs)
                    if all(isinstance(v, (list, dict)) for v in data.values()):
                        # Multiple columns with array/dict values
                        df = pd.DataFrame(data)
                    else:
                        # Single record
                        df = pd.DataFrame([data])
            else:
                raise ValueError("Unsupported JSON structure")
            
            print(f"✅ JSON loading successful! Shape: {df.shape}")
            return df
            
        except Exception as e:
            raise Exception(f"Failed to load JSON file: {e}")


class CSVDataLoader(BaseDataLoader):
    """Data loader for CSV files (.csv)"""
    
    def get_supported_extensions(self) -> List[str]:
        return ['.csv']
    
    def load_data(self) -> pd.DataFrame:
        """Load CSV data with encoding detection"""
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                df = pd.read_csv(self.file_path, encoding=encoding)
                print(f"✅ CSV loading successful with {encoding} encoding! Shape: {df.shape}")
                return df
            except Exception as e:
                print(f"❌ Failed with {encoding} encoding: {e}")
                continue
        
        raise Exception(f"Failed to load CSV with any encoding")


class ParquetDataLoader(BaseDataLoader):
    """Data loader for Parquet files (.parquet)"""
    
    def get_supported_extensions(self) -> List[str]:
        return ['.parquet']
    
    def load_data(self) -> pd.DataFrame:
        """Load Parquet data"""
        try:
            df = pd.read_parquet(self.file_path)
            print(f"✅ Parquet loading successful! Shape: {df.shape}")
            return df
        except Exception as e:
            raise Exception(f"Failed to load Parquet file: {e}")


def convert_csv_to_xlsx(csv_path: str, xlsx_path: str = None) -> str:
    """
    Convert CSV file to XLSX format.
    
    Args:
        csv_path: Path to input CSV file
        xlsx_path: Path to output XLSX file (if None, creates in same directory)
        
    Returns:
        Path to the converted XLSX file
    """
    try:
        # Read CSV file
        print(f"📄 Reading CSV file: {csv_path}")
        df = pd.read_csv(csv_path)
        
        # Generate output path if not provided
        if xlsx_path is None:
            csv_path_obj = Path(csv_path)
            xlsx_path = csv_path_obj.parent / f"{csv_path_obj.stem}.xlsx"
        
        # Write to XLSX
        print(f"💾 Converting to XLSX: {xlsx_path}")
        df.to_excel(xlsx_path, index=False, engine='openpyxl')
        
        print(f"✅ CSV to XLSX conversion successful!")
        print(f"   Input:  {csv_path} ({df.shape[0]} rows, {df.shape[1]} columns)")
        print(f"   Output: {xlsx_path}")
        
        return str(xlsx_path)
        
    except Exception as e:
        print(f"❌ Error converting CSV to XLSX: {e}")
        raise


class DataLoaderFactory:
    """Factory class to create appropriate data loader based on file extension"""
    
    _loaders = {
        '.xlsx': ExcelDataLoader,
        '.xls': ExcelDataLoader,
        '.json': JSONDataLoader,
        '.csv': CSVDataLoader,
        '.parquet': ParquetDataLoader,
    }
    
    @classmethod
    def create_loader(cls, file_path: str) -> BaseDataLoader:
        """Create appropriate data loader based on file extension"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext not in cls._loaders:
            supported_formats = list(cls._loaders.keys())
            raise ValueError(f"Unsupported file format: {file_ext}. Supported formats: {supported_formats}")
        
        loader_class = cls._loaders[file_ext]
        return loader_class(file_path)
    
    @classmethod
    def get_supported_formats(cls) -> List[str]:
        """Get list of all supported file formats"""
        return list(cls._loaders.keys())
    
    @classmethod
    def register_loader(cls, extension: str, loader_class: type):
        """Register a new data loader for a specific extension"""
        if not issubclass(loader_class, BaseDataLoader):
            raise ValueError("Loader class must inherit from BaseDataLoader")
        cls._loaders[extension] = loader_class


# Example of how to add a custom loader
class XMLDataLoader(BaseDataLoader):
    """Example: Data loader for XML files (.xml)"""
    
    def get_supported_extensions(self) -> List[str]:
        return ['.xml']
    
    def load_data(self) -> pd.DataFrame:
        """Load XML data (requires xml parsing logic)"""
        # Implementation would go here
        raise NotImplementedError("XML loader not implemented yet")


# Register the XML loader (example)
# DataLoaderFactory.register_loader('.xml', XMLDataLoader)


def load_data_from_file(file_path: str) -> pd.DataFrame:
    """Convenience function to load data from any supported file format"""
    loader = DataLoaderFactory.create_loader(file_path)
    return loader.load_data()


if __name__ == "__main__":
    # Example usage
    print("Supported formats:", DataLoaderFactory.get_supported_formats())
    
    # Example loading
    # df = load_data_from_file("/path/to/your/file.xlsx")
    # df = load_data_from_file("/path/to/your/file.json")
    # df = load_data_from_file("/path/to/your/file.csv")
