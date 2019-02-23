from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formula.translate import Translator


# TODO: Calculate these numbers in dynamically
COLUMN_LENGTH = 6 # The length of each day/slot, determines overall alignment
BEGIN_COLUMN = 2 #  We start in the 2nd column i.e. B for each day/slot
BEGIN_FREQ_ROW = 4 # We start at row 4 for each day/slot
BEGIN_SLOT_ROW = 6
NEXT_SLOT_ROW = 20
NEXT_COLUMN = COLUMN_LENGTH + 2 # Where the next column begins for each day/slot
# TODO: Move this into a style module
COLOR_LIGHTBLACK='00282828'
COLOR_DARKGREY='00505050'
COLOR_DARKRED='00600000'
# TODO: Make these user defineable
VOLUME_HEADERS = [ "Sets", "Load", "Reps", "RIR", "RPE",  "Avg Vel", "Int %" ]

# TODO: Move this into a style module
ALIGNMENT = Alignment(
    wrap_text=True, horizontal="center", vertical="center"
)


class Workout:

  def __init__(self, weeks=8, frequency=3, slots=3, sets=10):
    self.wb = Workbook()
    self.weeks = weeks     # How many weeks for the progrqm
    self.frequency = frequency # How many days per week
    self.slots = slots     # How many slots per day
    self.sets = sets     # How many sets per slot


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
          begin_col = BEGIN_COLUMN
          currentSheet = self.wb[sheet]

          for day in range(1, frequency + 1):
              # Add day header e.g .[ Day 1 ] [ Day 2 ] [ Day 3 ]

              currentCell = self.generate_header(
                  begin_row, begin_col, currentSheet, heading='Day', value=day
              )

              self.set_style(
                  currentSheet, currentCell, begin_col,
                  fgColor=colors.WHITE, bgColor=COLOR_LIGHTBLACK,
                  size=42, width=20, font='Helvetica', bold=True
              )

              begin_col += NEXT_COLUMN

      return frequency


  def generate_slots(self, slots: int, sets: int, frequency: int) -> int:
      # Add exercise slots to chosen frequency e.g.
      if not slots:
          # Set default
          slots = self.slots

      if not frequency:
          # Set default
          frequency = self.frequency

      if not sets:
          # Set default
          sets = self.sets

      for sheet in self.wb.sheetnames:
          # Get sheet
          # Generate tables for days in sheet
          currentSheet = self.wb[sheet]

          slot_col = BEGIN_COLUMN

          for day in range(1, frequency + 1):

              # TODO: Determining placement can be done better than this
              slot_row = BEGIN_SLOT_ROW
              exercise_row = BEGIN_SLOT_ROW + 2
              programming_row = BEGIN_SLOT_ROW + 3
              notes_row = BEGIN_SLOT_ROW + 4
              volume_header_row = BEGIN_SLOT_ROW + 5
              volume_input_row = BEGIN_SLOT_ROW + 6
              averages_row = volume_input_row + sets
              sums_row = volume_input_row + sets + 1

              for slot in range(1, slots + 1):
                  #print(f"Writing {sheet} row: {slot_row}, col: {slot_col}")

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
                      fgColor=colors.WHITE, bgColor=COLOR_DARKGREY,
                      size=32, width=20, font='Helvetica'
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
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=18, width=20, font='Helvetica', bold=False
                  )
                  exercise_row += NEXT_SLOT_ROW

                  # [       Program      ]
                  # [        Notes       ]
                  self.generate_divide(programming_row, slot_col, currentSheet, heading='Program')
                  self.generate_divide(notes_row, slot_col, currentSheet, heading='Notes')
                  # TODO: Row height can be set in a better place
                  currentSheet.row_dimensions[programming_row].height = 40
                  currentSheet.row_dimensions[notes_row].height = 40
                  programming_row += NEXT_SLOT_ROW
                  notes_row += NEXT_SLOT_ROW

                  # Add set header inputs
                  # [ Sets ] [ Load ] [ Reps ] [ RIR ] [ RPE ] [ Avg Vel ] [ Intensity ]
                  self.generate_volume_header(volume_header_row, slot_col, currentSheet)
                  volume_header_row += NEXT_SLOT_ROW

                  # Add set inputs
                  # [ Set 1 ] [ <input> ]
                  # [ Set 2 ] [ <input> ]
                  self.generate_volume_input(volume_input_row, slot_col, currentSheet, sets=sets)
                  # TODO: We should not be be referencing numbers, it's barely readable
                  self.generate_rir_to_rpe(volume_input_row, slot_col+4, currentSheet, sets=sets)
                  volume_input_row += NEXT_SLOT_ROW


                  # Add averages row
                  # [ Avgs ] [ <formula> ], etc.
                  self.generate_averages_row(averages_row, slot_col, currentSheet, sets=sets)
                  averages_row += NEXT_SLOT_ROW

                  # Add averages row
                  # [ Sums ] [ <formula> ], etc.
                  self.generate_sums_row(sums_row, slot_col, currentSheet, sets=sets)
                  sums_row += NEXT_SLOT_ROW

              # Start writing in column for next day
              slot_col += NEXT_COLUMN
      return slots


  def generate_header(self, row: int, col: int, currentSheet: object, heading: str = 'Header', value: str = 'Item') -> object:
              # Add horizontal header
              # [ Day 1 ]
              currentSheet.merge_cells(
                  start_row=row, end_row=row,
                  start_column=col, end_column=col+COLUMN_LENGTH
              )

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading} {value}"
              )

              return currentCell


  def generate_divide(self, row: int, col: int, currentSheet: object, heading: str = 'Header') -> object:
              # Create divide with header and input
              # [         ][         ]
              # [ Program ][ <input> ]
              # [         ][         ]

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading}"
              )

              currentSheet.merge_cells(
                  start_row=row, end_row=row, start_column=col+1, end_column=col+COLUMN_LENGTH
              )

              self.set_style(
                  currentSheet, currentCell, col,
                  fgColor=colors.WHITE, bgColor=COLOR_DARKGREY,
                  size=12, width=20, font='Helvetica', bold=False
              )

              currentCell = currentSheet.cell(
                  row=row, column=col+1
              )

              currentCell.alignment = ALIGNMENT

              return currentCell


  def generate_volume_header(self, row: int, col: int, currentSheet: object) -> object:
              for header in VOLUME_HEADERS:
                  currentCell = currentSheet.cell(
                      row=row, column=col, value=f"{header}"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=12, width=15, font='Helvetica', bold=True
                  )
                  # Set next column
                  col += 1

              return currentCell


  def generate_volume_input(self, row: int, col: int, currentSheet: object, sets: int) -> object:

              number_of_inputs = 5

              for number in range(1, sets + 1):

                  currentCell = currentSheet.cell(
                      row=row, column=col, value=f"Set {number}"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKGREY,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  for item in range(1, number_of_inputs + 1):

                      currentCell = currentSheet.cell(
                          row=row, column=col+item, value=""
                      )

                      currentCell.alignment = ALIGNMENT

                  # Set next column
                  row += 1

              return currentCell

  def generate_rir_to_rpe(self, row: int, col: int, currentSheet: object, sets: int) -> object:
              # [ RIR ] to [ RPE ]
              # [  2  ]    [  8  ]
              # e.g. #IF(E12="", "...", ABS(IFERROR(E12−10,"")))

              # The column before contains RPE
              col_rpe_letter = get_column_letter(col-1)

              for input_row in range(row, row + sets):

                  currentCell = currentSheet.cell(
                      row=row, column=col,
                      value=f"=IF({col_rpe_letter}{input_row}=\"\", \"...\", ABS(IFERROR({col_rpe_letter}{input_row}−10, \"\")))"
                  )

                  currentCell.alignment = ALIGNMENT

                  row += 1

              return currentCell


  def generate_averages_row(self, row: int, col: int, currentSheet: object, sets: int) -> object:

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"Averages"
              )

              self.set_style(
                  currentSheet, currentCell, col,
                  fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                  size=12, width=15, font='Helvetica', bold=True
              )

              col += 1

              # Get first row of user inputs [ Load ] [ Reps ], etc.
              begin_input_row = row - sets
              # Get last input row [ Load ] [ Reps ], etc.
              end_input_row = row - 1

              for input_row in range(begin_input_row, begin_input_row + sets):

                  col_letter = get_column_letter(col)

                  currentCell = currentSheet.cell(
                      row=row, column=col,
                      value=f"=IFERROR(ROUND(AVERAGE({col_letter}{begin_input_row}:{col_letter}{end_input_row}), 0), \"...\")"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  currentCell.alignment = ALIGNMENT

                  # Set next column
                  col += 1

                  if col == NEXT_COLUMN + 1:
                      break


  def generate_sums_row(self, row: int, col: int, currentSheet: object, sets: int) -> object:

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"Sums"
              )

              self.set_style(
                  currentSheet, currentCell, col,
                  fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                  size=12, width=15, font='Helvetica', bold=True
              )

              col += 1

              # Get first row of user inputs [ Load ] [ Reps ], etc.
              begin_input_row = row - sets
              # Get last input row [ Load ] [ Reps ], etc.
              end_input_row = row - 1

              for input_row in range(begin_input_row, begin_input_row + sets):

                  col_letter = get_column_letter(col)

                  currentCell = currentSheet.cell(
                      row=row, column=col,
                      value=f"=IFERROR(SUM({col_letter}{begin_input_row}:{col_letter}{end_input_row}), \"...\")"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  currentCell.alignment = ALIGNMENT

                  # Set next column
                  col += 1

                  if col == NEXT_COLUMN + 1:
                      break


  def set_style(self, sheet: object, cell: object, col: int, fgColor: str, bgColor: str, size: int, width: int, font: str, bold: bool = False) -> object:
        # Set style
        font = Font(
            name=font, size=size, bold=bold, color=fgColor
        )
        fill = PatternFill(
            fill_type='solid', fgColor=bgColor,
        )

        sheet.column_dimensions[get_column_letter(col)].width = width
        cell.font = font
        cell.fill = fill
        cell.alignment = ALIGNMENT
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
