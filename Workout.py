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
      for week in range(1, weeks + 1):
         # Create a new sheet for each week
         sheet=f"Week {week}"
         print(f"Writing sheet {sheet}")
         ws = self.wb.create_sheet(title=sheet)

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
          currentSheet = self.wb[sheet]

          for day in range(1, frequency + 1):
              # Add day header e.g .[ Day 1 ] [ Day 2 ] [ Day 3 ]

              currentCell = self.generate_header(
                  begin_row, begin_col, currentSheet, heading='Day', value=day
              )

              self.set_style(
                  currentSheet, currentCell, begin_col,
                  color=colors.BLACK, size=42, width=20, font='Helvetica'
              )

              begin_col = begin_col + 5

      return frequency


  def generate_slots(self, slots: int, frequency: int) -> int:
      # Add exercise slots to chosen frequency e.g.
      # [    Day 1   ]
      # [ Exercise 1 ]
      # [ Exercise 2 ]
      # [ Exercise 3 ]

      if not slots:
          # Set default
          slots = self.slots

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          currentSheet = self.wb[sheet]

          begin_col = 4 # We start in the 4th column i.e. D

          for day in range(1, frequency + 1):

              begin_row = 6 # We start in the 6th row

              for slot in range(1, slots + 1):
                  print(f"Writing {sheet} row: {begin_row}, col: {begin_col}")

                  currentCell = self.generate_header(
                      begin_row, begin_col, currentSheet, heading='Exercise', value=slot
                  )

                  self.set_style(
                      currentSheet, currentCell, begin_col,
                      color=colors.RED, size=32, width=20, font='Helvetica'
                  )

                  begin_row += + 20

              # Start writing in column for next day
              begin_col += 5

      return slots


  def generate_header(self, row: int, col: int, currentSheet: object, heading: str = 'Header', value: str = 'Item') -> object:
              currentSheet.merge_cells(
                  start_row=row, end_row=row,
                  start_column=col, end_column=col+3
              )

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading} {value}"
              )

              return currentCell

  def save(self, filename: str) -> str:
      if not filename:
          # Set default
          filename = 'workout.xlsx'

      self.wb.save(filename)
      print(f"Writing program to {filename}")
      return filename


  def set_style(self, sheet: object, cell: object, col: int, color: str, size: int, width: int, font: str) -> object:
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

        sheet.column_dimensions[get_column_letter(col)].width = width
        cell.font = font
        cell.fill = fill
        cell.alignment = alignment
        return cell


  def test(self, msg: str) -> str:
      if not msg:
          msg = 'Test'
      return msg
