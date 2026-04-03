from pathlib import Path
import pandas as pd
import sys

class ExcelMaster:
    def __init__(self):
        self.df = None
        self.path = Path(__file__).parent.resolve()
        self.translate = {
            
                
        }
        pass
    #Union de todos los excel     
    def merge_files(self, extension=".xlsx"):
        
        files = list(self.path.glob(f"*{extension}"))

        if not files:
            print(f"❌ No files found in {self.path}")
            return
        
        frames = []

        for file in files:
            # --- 1. BARRA DE PROGRESO / LOG DE ERRORES ---
            try:
                if file.suffix == ".csv":
                    df_temp = pd.read_csv(file)
                else:
                    # --- 2. MANEJO DE VARIAS HOJAS (Sheets) ---
                    # sheet_name=None lee TODAS las hojas y devuelve un diccionario
                    dict_sheets = pd.read_excel(file, sheet_name=None)
                    # Unimos todas las hojas de ese archivo en un solo dataframe temporal
                    df_temp = pd.concat(dict_sheets.values(), ignore_index=True)

                # --- 3. UNIFICADOR DE COLUMNAS (The Column Fixer) ---
                # Ponemos todo en minúsculas para que 'Price' y 'price' sean lo mismo
                df_temp.columns = [str(c).lower().strip() for c in df_temp.columns]
                
                frames.append(df_temp)
                print(f"✅ Loaded: {file.name}")

            except Exception as e:
                # Si un archivo está abierto o corrupto, el bot no se detiene
                print(f"⚠️ Skipping {file.name}: Error -> {e}")

        if frames:
            self.df = pd.concat(frames, ignore_index=True)
            print(f"🚀 Successfully merged {len(frames)} files.")

    #calculos
    def calculate_total (self, qty_col, price_col, target_col="Total"):
        """
        Standard Tier Feature: Custom math calculations
        Example: df['Total'] = df['Quantity'] * df['Price']
        """
        if self.df is not None:
            self.df[target_col] = self.df[qty_col] * self.df[price_col]
            print(f"📊 Column '{target_col}' calculated successfully.")
    
    
    #custom calculo USD
    def price_to_USD (self, variable_columna, target_USD="Price_USD"):
        self.df[target_USD] = self.df[variable_columna] / 4000
    
    #hacer clean en el excel 
    def smart_clean(self):
        """
        Removes duplicates and empty rows
        """
        if self.df is not None:
            before = len(self.df)
            self.df.drop_duplicates(inplace=True)
            self.df.dropna(how='all', inplace=True)
            print(f"🧹 Cleaned: {before - len(self.df)} redundant rows removed.")
    
    #guardar
    def save_report(self, folder_name="Results", filename="final_report.xlsx"):
        output_dir = self.path / folder_name
        output_dir.mkdir(parents=True, exist_ok=True)
        final_file = output_dir / filename

        # Usamos XlsxWriter para que se vea PRO
        with pd.ExcelWriter(final_file, engine='xlsxwriter') as writer:
            self.df.to_excel(writer, index=False, sheet_name='Data')
            
            workbook  = writer.book
            worksheet = writer.sheets['Data']

            # Formato: Encabezado Azul y Letra Blanca
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1
            })

            # Aplicar formato a encabezados
            for col_num, value in enumerate(self.df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                # Ajustar ancho de columna automáticamente (aproximado)
                worksheet.set_column(col_num, col_num, 20)

        print(f"💾 File exported with STYLE to: {final_file}")


    #crear excel
    def create_empty_report(self, columns_list):
        """
        Creates a brand new blank DataFrame with specific columns.
        Useful for starting a project from scratch.
        """
        self.df = pd.DataFrame(columns=columns_list)
        print(f"✨ New blank report created with columns: {columns_list}")
    
    #agregar al excel
    def add_row(self, data_dict):
        """
        Adds a single row of data to the current report.
        'data_dict' should be like {'Name': 'Eduardo', 'Price': 4000}
        """
        if self.df is not None:
            # Creamos un DF temporal con la nueva fila y lo concatenamos
            new_row = pd.DataFrame([data_dict])
            self.df = pd.concat([self.df, new_row], ignore_index=True)
            print(f"➕ Row added: {list(data_dict.values())[0]}...")

    

#iniciador 
if __name__ == "__main__":
    bot = ExcelMaster()
    
    # 1. Unir todo lo que sea .xlsx en la carpeta
    bot.merge_files(extension=".xlsx")
    
    # 2. Si el cliente pidió cálculos (Nivel Standard)
    # bot.calculate_total("Quantity", "Price")
    
    # 3. Limpiar la data
    bot.smart_clean()
    #custom
    bot.price_to_USD("precio")
    
    # 4. Guardar
    bot.save_report(filename="Client_Result.xlsx")


            

