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

      # Set style
      font = Font(
          name='Calibri', size=48, bold=True, color=colors.WHITE
      )
      fill = PatternFill(
          fill_type='solid', bgColor=colors.BLACK,
      )
      alignment = Alignment(
          horizontal="center", vertical="center"
      )

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          begin_row = 4 # We start in the 4th row
          begin_col = 4 # We start in the 4th column i.e. D
          for day in range(1, frequency + 1):
              print(day)
              # Add day header e.g .[ Day 1 ] [ Day 2 ] [ Day 3 ]
              cws = self.wb[sheet]
              cws.column_dimensions[get_column_letter(begin_col)].width = 60
              cws.cell(row=begin_row, column=begin_col, value=f"Day {day}")
              cws.fill = fill
              cws.alignment = alignment
              begin_col = begin_col + 2

      return frequency


  def generate_slots(self, slots: int) -> int:
      if not slots:
          # Set default
          slots = self.slots

      return slots

  def save(self, filename: str) -> str:
      if not filename:
          # Set default
          filename = 'workout.xlsx'

      self.wb.save(filename)
      return filename
