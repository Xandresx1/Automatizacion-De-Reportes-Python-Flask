"""
Módulo de procesamiento y limpieza de datos para la herramienta de automatización de reportes.

Este módulo se encarga de cargar, limpiar y preparar los datos para la generación de reportes.
Incluye funciones para eliminar duplicados, normalizar fechas, y validar la estructura de los datos.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import logging
from typing import Tuple, List, Dict, Optional
import os

# Configurar logging para debug
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DataProcessor:
    """
    Clase principal para el procesamiento de datos.
    
    Esta clase maneja toda la lógica de limpieza y preparación de datos,
    incluyendo la detección automática de tipos de datos y la normalización
    de formatos.
    """
    
    def __init__(self):
        """Inicializa el procesador de datos."""
        self.original_data = None
        self.cleaned_data = None
        self.data_info = {}
        
    def load_data(self, file_path: str) -> Tuple[bool, str]:
        """
        Carga datos desde un archivo CSV o Excel.
        
        Args:
            file_path: Ruta al archivo de datos
            
        Returns:
            Tuple con (éxito, mensaje)
        """
        try:
            # Detectar el tipo de archivo y cargarlo apropiadamente
            if file_path.endswith('.csv'):
                self.original_data = pd.read_csv(file_path)
                logger.info(f"Archivo CSV cargado exitosamente: {file_path}")
            elif file_path.endswith(('.xlsx', '.xls')):
                self.original_data = pd.read_excel(file_path)
                logger.info(f"Archivo Excel cargado exitosamente: {file_path}")
            else:
                return False, "Formato de archivo no soportado. Usa CSV o Excel."
            
            # Verificar que el archivo no esté vacío
            if self.original_data.empty:
                return False, "El archivo está vacío o no contiene datos válidos."
            
            # Guardar información básica del dataset
            self.data_info = {
                'filas_originales': len(self.original_data),
                'columnas': list(self.original_data.columns),
                'tipos_datos': self.original_data.dtypes.to_dict()
            }
            
            return True, f"Datos cargados exitosamente. {len(self.original_data)} filas, {len(self.original_data.columns)} columnas."
            
        except Exception as e:
            logger.error(f"Error al cargar datos: {str(e)}")
            return False, f"Error al cargar el archivo: {str(e)}"
    
    def clean_data(self) -> Tuple[bool, str]:
        """
        Limpia y prepara los datos para el análisis.
        
        Este método realiza las siguientes operaciones de limpieza:
        - Elimina filas completamente vacías
        - Elimina duplicados (que pueden falsear los resultados)
        - Normaliza fechas para evitar problemas posteriores
        - Convierte tipos de datos apropiadamente
        - Maneja valores faltantes
        
        Returns:
            Tuple con (éxito, mensaje)
        """
        if self.original_data is None:
            return False, "No hay datos cargados para limpiar."
        
        try:
            # Hacer una copia para no modificar los datos originales
            self.cleaned_data = self.original_data.copy()
            
            # 1. Eliminar filas completamente vacías
            filas_antes = len(self.cleaned_data)
            self.cleaned_data = self.cleaned_data.dropna(how='all')
            filas_eliminadas = filas_antes - len(self.cleaned_data)
            if filas_eliminadas > 0:
                logger.info(f"Eliminadas {filas_eliminadas} filas completamente vacías")
            
            # 2. Eliminar duplicados - esto es crucial para evitar sesgos en los reportes
            duplicados_antes = len(self.cleaned_data)
            self.cleaned_data = self.cleaned_data.drop_duplicates()
            duplicados_eliminados = duplicados_antes - len(self.cleaned_data)
            if duplicados_eliminados > 0:
                logger.info(f"Eliminados {duplicados_eliminados} registros duplicados")
            
            # 3. Normalizar fechas - esto evita problemas con diferentes formatos
            self._normalize_dates()
            
            # 4. Convertir tipos de datos apropiadamente
            self._convert_data_types()
            
            # 5. Manejar valores faltantes
            self._handle_missing_values()
            
            # Actualizar información del dataset
            self.data_info.update({
                'filas_finales': len(self.cleaned_data),
                'filas_eliminadas': filas_antes - len(self.cleaned_data),
                'columnas_numericas': self._get_numeric_columns(),
                'columnas_fecha': self._get_date_columns()
            })
            
            return True, f"Datos limpiados exitosamente. {len(self.cleaned_data)} filas finales."
            
        except Exception as e:
            logger.error(f"Error durante la limpieza: {str(e)}")
            return False, f"Error durante la limpieza: {str(e)}"
    
    def _normalize_dates(self):
        """
        Normaliza las columnas de fechas para asegurar consistencia.
        
        Esta función busca columnas que parezcan fechas y las convierte
        al formato estándar de pandas. Esto es importante porque diferentes
        fuentes de datos pueden tener formatos de fecha diferentes.
        """
        for col in self.cleaned_data.columns:
            # Intentar detectar si la columna contiene fechas
            if self._is_date_column(col):
                try:
                    # Convertir a datetime y manejar errores
                    self.cleaned_data[col] = pd.to_datetime(
                        self.cleaned_data[col], 
                        errors='coerce',  # Los valores que no se pueden convertir se vuelven NaT
                        infer_datetime_format=True
                    )
                    logger.info(f"Columna de fecha normalizada: {col}")
                except Exception as e:
                    logger.warning(f"No se pudo normalizar la columna {col} como fecha: {e}")
    
    def _is_date_column(self, column_name: str) -> bool:
        """
        Detecta si una columna probablemente contiene fechas.
        
        Args:
            column_name: Nombre de la columna a verificar
            
        Returns:
            True si la columna parece contener fechas
        """
        # Palabras clave que sugieren que es una columna de fecha
        date_keywords = ['fecha', 'date', 'time', 'tiempo', 'dia', 'day', 'mes', 'month', 'año', 'year']
        
        # Verificar si el nombre de la columna contiene palabras clave de fecha
        col_lower = column_name.lower()
        if any(keyword in col_lower for keyword in date_keywords):
            return True
        
        # Verificar si los primeros valores no nulos parecen fechas
        sample_values = self.cleaned_data[column_name].dropna().head(10)
        if len(sample_values) == 0:
            return False
        
        # Intentar convertir algunos valores de muestra
        try:
            pd.to_datetime(sample_values, errors='raise')
            return True
        except:
            return False
    
    def _convert_data_types(self):
        """
        Convierte tipos de datos apropiadamente para optimizar el análisis.
        
        Esta función identifica columnas numéricas y las convierte al tipo
        apropiado (int, float) para facilitar los cálculos posteriores.
        """
        for col in self.cleaned_data.columns:
            # Saltar columnas de fecha que ya fueron procesadas
            if self.cleaned_data[col].dtype == 'datetime64[ns]':
                continue
            
            # Intentar convertir a numérico si es apropiado
            if self._is_numeric_column(col):
                try:
                    # Intentar convertir a float primero
                    self.cleaned_data[col] = pd.to_numeric(self.cleaned_data[col], errors='coerce')
                    
                    # Si todos los valores son enteros, convertir a int
                    if self.cleaned_data[col].dropna().apply(float.is_integer).all():
                        self.cleaned_data[col] = self.cleaned_data[col].astype('Int64')  # Int64 maneja NaN
                    
                    logger.info(f"Columna numérica convertida: {col}")
                except Exception as e:
                    logger.warning(f"No se pudo convertir la columna {col} a numérico: {e}")
    
    def _is_numeric_column(self, column_name: str) -> bool:
        """
        Detecta si una columna probablemente contiene valores numéricos.
        
        Args:
            column_name: Nombre de la columna a verificar
            
        Returns:
            True si la columna parece contener valores numéricos
        """
        # Palabras clave que sugieren que es una columna numérica
        numeric_keywords = ['cantidad', 'total', 'suma', 'precio', 'valor', 'numero', 'number', 'count', 'amount', 'sales', 'ventas']
        
        col_lower = column_name.lower()
        if any(keyword in col_lower for keyword in numeric_keywords):
            return True
        
        # Verificar si los primeros valores no nulos parecen números
        sample_values = self.cleaned_data[column_name].dropna().head(10)
        if len(sample_values) == 0:
            return False
        
        # Intentar convertir algunos valores de muestra
        try:
            pd.to_numeric(sample_values, errors='raise')
            return True
        except:
            return False
    
    def _handle_missing_values(self):
        """
        Maneja valores faltantes de manera apropiada según el tipo de dato.
        
        Para columnas numéricas, reemplaza NaN con 0 o la media.
        Para columnas de texto, reemplaza NaN con "No especificado".
        Para fechas, deja NaN ya que no tiene sentido inventar fechas.
        """
        for col in self.cleaned_data.columns:
            if self.cleaned_data[col].isnull().sum() > 0:
                if self.cleaned_data[col].dtype in ['int64', 'float64', 'Int64']:
                    # Para columnas numéricas, usar 0 o la media
                    if self.cleaned_data[col].dtype == 'float64':
                        self.cleaned_data[col].fillna(0, inplace=True)
                    else:
                        self.cleaned_data[col].fillna(0, inplace=True)
                    logger.info(f"Valores faltantes en columna numérica '{col}' reemplazados con 0")
                
                elif self.cleaned_data[col].dtype == 'object':
                    # Para columnas de texto, usar "No especificado"
                    self.cleaned_data[col].fillna("No especificado", inplace=True)
                    logger.info(f"Valores faltantes en columna de texto '{col}' reemplazados con 'No especificado'")
                
                # Para fechas, no hacemos nada - NaN es apropiado
    
    def _get_numeric_columns(self) -> List[str]:
        """Obtiene la lista de columnas numéricas."""
        numeric_cols = []
        for col in self.cleaned_data.columns:
            if self.cleaned_data[col].dtype in ['int64', 'float64', 'Int64']:
                numeric_cols.append(col)
        return numeric_cols
    
    def _get_date_columns(self) -> List[str]:
        """Obtiene la lista de columnas de fecha."""
        date_cols = []
        for col in self.cleaned_data.columns:
            if self.cleaned_data[col].dtype == 'datetime64[ns]':
                date_cols.append(col)
        return date_cols
    
    def get_data_summary(self) -> Dict:
        """
        Genera un resumen de los datos procesados.
        
        Returns:
            Diccionario con información resumida de los datos
        """
        if self.cleaned_data is None:
            return {"error": "No hay datos procesados"}
        
        summary = {
            "filas_originales": self.data_info.get('filas_originales', 0),
            "filas_finales": self.data_info.get('filas_finales', 0),
            "filas_eliminadas": self.data_info.get('filas_eliminadas', 0),
            "columnas_totales": len(self.cleaned_data.columns),
            "columnas_numericas": len(self._get_numeric_columns()),
            "columnas_fecha": len(self._get_date_columns()),
            "columnas_texto": len(self.cleaned_data.columns) - len(self._get_numeric_columns()) - len(self._get_date_columns()),
            "valores_faltantes": self.cleaned_data.isnull().sum().sum(),
            "columnas": list(self.cleaned_data.columns)
        }
        
        return summary
    
    def get_sample_data(self, n_rows: int = 5) -> pd.DataFrame:
        """
        Obtiene una muestra de los datos procesados.
        
        Args:
            n_rows: Número de filas a mostrar
            
        Returns:
            DataFrame con la muestra de datos
        """
        if self.cleaned_data is None:
            return pd.DataFrame()
        
        return self.cleaned_data.head(n_rows)
    
    def get_cleaned_data(self) -> pd.DataFrame:
        """
        Obtiene los datos limpios procesados.
        
        Returns:
            DataFrame con los datos limpios
        """
        return self.cleaned_data.copy() if self.cleaned_data is not None else pd.DataFrame()
