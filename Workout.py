from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter, column_index_from_string

BEGIN_FREQ_ROW = 4
BEGIN_FREQ_COL = 2 # We start in the 2nd column i.e. B
BEGIN_SLOT_COLUMN = 2 # We start in the 2nd column i.e. B
BEGIN_SLOT_ROW = 6
NEXT_SLOT_ROW = 20
NEXT_SLOT_COLUMN = 7
NEXT_DAY_COLUMN = 7
NEXT_DIVIDE_COLUMN = 5
HEADER_LENGTH = 5

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
          begin_row = BEGIN_FREQ_ROW
          begin_col = BEGIN_FREQ_COL
          currentSheet = self.wb[sheet]

          for day in range(1, frequency + 1):
              # Add day header e.g .[ Day 1 ] [ Day 2 ] [ Day 3 ]

              currentCell = self.generate_header(
                  begin_row, begin_col, currentSheet, heading='Day', value=day
              )

              self.set_style(
                  currentSheet, currentCell, begin_col,
                  color=colors.BLACK, size=42, width=20, font='Helvetica', bold=True
              )

              begin_col += NEXT_DAY_COLUMN

      return frequency


  def generate_slots(self, slots: int, frequency: int) -> int:
      # Add exercise slots to chosen frequency e.g.
      if not slots:
          # Set default
          slots = self.slots

      if not frequency:
          # Set default
          frequency = self.frequency

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          currentSheet = self.wb[sheet]

          slot_col = BEGIN_SLOT_COLUMN

          for day in range(1, frequency + 1):

              slot_row = BEGIN_SLOT_ROW # We start in the 6th row
              exercise_row = 8
              programming_row = 9
              notes_row = 10

              for slot in range(1, slots + 1):
                  print(f"Writing {sheet} row: {slot_row}, col: {slot_col}")

                  # Add exercise slot header
                  # [    Day 1   ]
                  # [ Exercise 1 ]
                  # [ Exercise 2 ]
                  # [ Exercise 3 ]
                  currentCell = self.generate_header(
                      slot_row, slot_col, currentSheet, heading='Exercise', value=slot
                  )
                  self.set_style(
                      currentSheet, currentCell, slot_col,
                      color=colors.RED, size=32, width=20, font='Helvetica'
                  )
                  slot_row += NEXT_SLOT_ROW

                  # Add exercise header
                  # [    Day 1   ]
                  # [ Exercise 1 ]
                  # [  Exercise  ]
                  # [ Exercise 2 ]
                  # [  Exercise  ]
                  currentCell = self.generate_header(
                      exercise_row, slot_col, currentSheet, heading='Exercise', value=''
                  )
                  self.set_style(
                      currentSheet, currentCell, slot_col,
                      color='00808080', size=24, width=20, font='Helvetica', bold=True
                  )
                  exercise_row += NEXT_SLOT_ROW

                  self.generate_divide(programming_row, slot_col, currentSheet, heading='Volume & Intensity')
                  self.generate_divide(notes_row, slot_col, currentSheet, heading='Notes')

                  programming_row += NEXT_SLOT_ROW                    
                  notes_row += NEXT_SLOT_ROW                    
                  # [ Sets ] [ Weight ] [ Reps ] [ RIR ] [ RPE ] [ Intensity ]

              # Start writing in column for next day
              slot_col += NEXT_SLOT_COLUMN
      return slots


  def generate_header(self, row: int, col: int, currentSheet: object, heading: str = 'Header', value: str = 'Item') -> object:
              # Add horizontal header
              # [ Day 1 ]
              currentSheet.merge_cells(
                  start_row=row, end_row=row,
                  start_column=col, end_column=col+HEADER_LENGTH
              )

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading} {value}"
              )

              return currentCell


  def generate_divide(self, row: int, col: int, currentSheet: object, heading: str = 'Header') -> object:
              # Create divide with header and input
              # [        ][         ]
              # [ Volume ][ <input> ]
              # [        ][         ]

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading}"
              )

              currentSheet.merge_cells(
                  start_row=row, end_row=row, start_column=col+1, end_column=col+NEXT_DIVIDE_COLUMN
              )

              self.set_style(
                  currentSheet, currentCell, col,
                  color=colors.WHITE, size=12, width=20, font='Helvetica', bold=False
              )

              return currentCell


  def set_style(self, sheet: object, cell: object, col: int, color: str, size: int, width: int, font: str, bold: bool = False) -> object:
        # Set style
        font = Font(
            name=font, size=size, bold=bold, color=colors.WHITE
        )
        fill = PatternFill(
            fill_type='solid', bgColor=color,
        )
        alignment = Alignment(
            horizontal="center", vertical="center"
        )

        sheet.column_dimensions[get_column_letter(col)].width = width
        #sheet.column_dimensions[get_column_letter(col)].height = height
        cell.font = font
        cell.fill = fill
        cell.alignment = alignment
        return cell


  def save(self, filename: str) -> str:
      if not filename:
          # Set default
          filename = 'workout.xlsx'

      self.wb.save(filename)
      print(f"Writing program to {filename}")
      return filename


  def test(self, msg: str) -> str:
      if not msg:
          msg = 'Test'
      return msg
