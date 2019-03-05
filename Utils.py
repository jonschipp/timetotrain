from Style import Style
from openpyxl.styles import PatternFill


class Utils:


  @staticmethod
  def set_formula(currentCell: object, formula: str) -> None:
        currentCell.value = formula
        currentCell.alignment = Style.Settings.ALIGNMENT


  @staticmethod
  def clear(workbook: object) -> None:
     # Make remaining of our cells white and borderless
     color = PatternFill(fill_type='solid', fgColor=Style.Settings.LIGHTBLACK)
     for sheet in workbook.worksheets:
         for row in sheet:
             for currentCell in row:
                 if currentCell.alignment.horizontal == 'center':
                     # Skip the cells we created
                     continue
                 currentCell.fill = color


  @staticmethod
  def save(workbook: object, filename: str) -> str:
      if not filename:
          # Set default
          filename = 'workout.xlsx'

      workbook.save(filename)
      print(f"Writing program to {filename}")
      return filename
