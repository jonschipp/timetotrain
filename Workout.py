from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter, column_index_from_string


class Workout:

  def __init__(self, weeks=8, frequency=3, slots=3):
    self.wb = Workbook()
    self.weeks = weeks     # How many weeks for the progrqm
    self.frequency = frequency # How many days per week
    self.slots = slots     # How many slots per day


  def generate_weeks(self, weeks: int) -> list:
      if not weeks:
          # Set default
          weeks = self.weeks

      # Weeks
      for week in range(1, weeks):
         # Create a new sheet for each week
         ws = self.wb.create_sheet(title=f"Week {week}")

      # Remove default sheet
      del self.wb['Sheet']

      return self.wb.sheetnames


  def generate_frequency(self, frequency: int) -> int:
      if not frequency:
          # Set default
          frequency = self.frequency

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          begin_row = 4 # We start in the 4th row
          begin_col = 4 # We start in the 4th column i.e. D
          for day in range(1, frequency + 1):
              # Add day header e.g .[ Day 1 ] [ Day 2 ] [ Day 3 ]
              currentSheet = self.wb[sheet]
              currentCell = currentSheet.cell(
                  row=begin_row, column=begin_col, value=f"Day {day}"
              )
              self.test('Blah')
              self.set_style(currentSheet, currentCell, begin_col, colors.BLACK, 48, 'Helvetica')
              begin_col = begin_col + 2

      return frequency


  def generate_slots(self, slots: int) -> int:
      if not slots:
          # Set default
          slots = self.slots

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          begin_row = 6 # We start in the 4th row
          begin_col = 4 # We start in the 4th column i.e. D
          for slot in range(1, slots + 1):
              # Add exercise slots e.g.
              # [ Exercise 1 ]
              # [ Exercise 2 ]
              # [ Exercise 3 ]
              currentSheet = self.wb[sheet]
              currentCell = currentSheet.cell(
                  row=begin_row, column=begin_col, value=f"Exercise {slot}"
              )
              self.test('Blah')
              self.set_style(currentSheet, currentCell, begin_col, colors.BLACK, 32, 'Helvetica')
              begin_col = begin_col + 2

      return slots


  def save(self, filename: str) -> str:
      if not filename:
          # Set default
          filename = 'workout.xlsx'

      self.wb.save(filename)
      return filename


def set_style(self, sheet: object, cell: object, col: int, color: str, size: int, font: str) -> object:
      # Set style
      font = Font(
          name=font, size=size, bold=True, color=colors.WHITE
      )
      fill = PatternFill(
          fill_type='solid', bgColor=color,
      )
      alignment = Alignment(
          horizontal="center", vertical="center"
      )

      sheet.column_dimensions[get_column_letter(col)].width = 60
      cell.font = font
      cell.fill = fill
      cell.alignment = alignment
      return cell

def test(self, msg: str) -> str:
    if not msg:
        msg = 'Test'
    return msg
