from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formula.translate import Translator


# TODO: Calculate these numbers in dynamically
COLUMN_LENGTH = 6 # The length of each day/slot, determines overall alignment
BEGIN_COLUMN = 2 #  We start in the 2nd column i.e. B for each day/slot
BEGIN_FREQ_ROW = 4 # We start at row 4 for each day/slot
BEGIN_SLOT_ROW = 6 # The row where the exercise slot begins e.g. [ Exercise 1 ]
NEXT_SLOT_ROW = 22 # If we add more rows, we need to increase this by 1 for each added row
NEXT_COLUMN = COLUMN_LENGTH + 2 # Where the next column begins for each day/slot
# TODO: Make these user defineable
VOLUME_HEADERS = {
    "Sets":    { "ColumnNumber": BEGIN_COLUMN,     "ColumnLetter": get_column_letter(BEGIN_COLUMN)    },
    "Load":    { "ColumnNumber": BEGIN_COLUMN + 1, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 1)},
    "Reps":    { "ColumnNumber": BEGIN_COLUMN + 2, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 2)},
    "RIR":     { "ColumnNumber": BEGIN_COLUMN + 3, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 3)},
    "RPE":     { "ColumnNumber": BEGIN_COLUMN + 4, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 4)},
    "Avg Vel": { "ColumnNumber": BEGIN_COLUMN + 5, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 5)},
    "Int %":   { "ColumnNumber": BEGIN_COLUMN + 6, "ColumnLetter": get_column_letter(BEGIN_COLUMN + 6)}
}
VOLUME_LENGTH = len(VOLUME_HEADERS)

# TODO: Move this into a style module
COLOR_LIGHTBLACK='00282828'
COLOR_DARKGREY='00505050'
COLOR_DARKRED='00600000'
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
              volume_row = volume_input_row + sets + 2
              tonnage_row = volume_input_row + sets + 3
              e1rm_row = volume_input_row + sets + 4

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
                  self.generate_volume_input(volume_input_row, slot_col, currentSheet, sets=sets, e1rm_row=e1rm_row)
                  # TODO: We should not be be referencing numbers, it's barely readable
                  self.generate_rir_to_rpe(volume_input_row, slot_col+4, currentSheet, sets=sets)


                  # Add averages row
                  # [ Avgs ] [ <formula> ], etc.
                  self.generate_averages_row(averages_row, slot_col, currentSheet, sets=sets)
                  averages_row += NEXT_SLOT_ROW

                  # Add averages row
                  # [ Sums ] [ <formula> ], etc.
                  self.generate_sums_row(sums_row, slot_col, currentSheet, sets=sets)
                  # Add row for Volume (sets x reps) that reads the Reps sum - for convenience.
                  # Depends on value of sums_row before we increment it
                  self.set_formula(
                      currentCell=self.generate_divide(volume_row, slot_col, currentSheet, heading='Volume', style='formula'),
                      formula=f"={VOLUME_HEADERS['Reps']['ColumnLetter']}{sums_row}"
                  )

                  self.set_formula(
                      currentCell=self.generate_divide(tonnage_row, slot_col, currentSheet, heading='Tonnage', style='formula'),
                      formula=self.generate_tonnage_formula(volume_input_row, sets)
                  )
                  self.generate_divide(e1rm_row, slot_col, currentSheet, heading='E1RM', style='manual')

                  volume_input_row += NEXT_SLOT_ROW
                  sums_row += NEXT_SLOT_ROW
                  tonnage_row += NEXT_SLOT_ROW
                  volume_row += NEXT_SLOT_ROW
                  e1rm_row += NEXT_SLOT_ROW

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


  def generate_divide(self, row: int, col: int, currentSheet: object, heading: str = 'Header', style: str = 'manual') -> object:
              # Create divide with header and input
              # [         ][         ]
              # [ Program ][ <input> ]
              # [         ][         ]
              color = COLOR_DARKGREY
              bold = False

              if style == 'formula':
                  color = COLOR_DARKRED
                  bold = True

              currentCell = currentSheet.cell(
                  row=row, column=col, value=f"{heading}"
              )

              currentSheet.merge_cells(
                  start_row=row, end_row=row, start_column=col+1, end_column=col+COLUMN_LENGTH
              )

              self.set_style(
                  currentSheet, currentCell, col,
                  fgColor=colors.WHITE, bgColor=color,
                  size=12, width=20, font='Helvetica', bold=bold
              )

              currentCell = currentSheet.cell(
                  row=row, column=col+1
              )

              if style == 'formula':
                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=color,
                      size=12, width=20, font='Helvetica', bold=False
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


  def generate_volume_input(self, row: int, col: int, currentSheet: object, sets: int, **kwargs: dict) -> object:

              for number in range(1, sets + 1):

                  currentCell = currentSheet.cell(
                      row=row, column=col, value=f"Set {number}"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKGREY,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  for item in range(1, VOLUME_LENGTH):

                      currentCell = currentSheet.cell(
                          row=row, column=col+item, value=""
                      )

                      # Add intensity calculation based off E1RM
                      if col+item == VOLUME_HEADERS['Int %']['ColumnNumber']:
                          self.set_formula(
                              currentCell=currentCell,
                              formula=f"=IF(ISBLANK({VOLUME_HEADERS['Load']['ColumnLetter']}{row}), \"...\", {VOLUME_HEADERS['Load']['ColumnLetter']}{row}/{VOLUME_HEADERS['Load']['ColumnLetter']}{kwargs['e1rm_row']})"
                          )
                          currentCell.number_format = '0%'

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

              count = 1

              for input_row in range(begin_input_row, begin_input_row + sets):

                  col_letter = get_column_letter(col)

                  currentCell = currentSheet.cell(
                      row=row, column=col,
                      value=f"=IFERROR(ROUND(AVERAGE({col_letter}{begin_input_row}:{col_letter}{end_input_row}), 0), \"...\")"
                  )

                  if col == VOLUME_HEADERS['Avg Vel']['ColumnNumber']:
                      currentCell.value = f"=IFERROR(AVERAGE({col_letter}{begin_input_row}:{col_letter}{end_input_row}), \"...\")"

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  currentCell.alignment = ALIGNMENT

                  if col == VOLUME_HEADERS['Int %']['ColumnNumber']:
                      currentCell.number_format = '0%'

                  # Set next column
                  col += 1
                  count += 1

                  if count == VOLUME_LENGTH:
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
              begin_input_row = row - sets - 1
              # Get last input row [ Load ] [ Reps ], etc.
              end_input_row = row - 2
              count = 1

              for input_row in range(begin_input_row, begin_input_row + sets):

                  col_letter = get_column_letter(col)

                  currentCell = currentSheet.cell(
                      row=row, column=col,
                      value=f"=IF(SUM({col_letter}{begin_input_row}:{col_letter}{end_input_row})>0, SUM({col_letter}{begin_input_row}:{col_letter}{end_input_row}), \"...\")"
                  )

                  self.set_style(
                      currentSheet, currentCell, col,
                      fgColor=colors.WHITE, bgColor=COLOR_DARKRED,
                      size=12, width=8, font='Helvetica', bold=False
                  )

                  currentCell.alignment = ALIGNMENT

                  if col == VOLUME_HEADERS['Int %']['ColumnNumber']:
                      currentCell.number_format = '0%'

                  # Set next column
                  col += 1
                  count += 1

                  if count == VOLUME_LENGTH:
                      break


  def set_formula(self, currentCell: object, formula: str) -> None:
        currentCell.value = formula
        currentCell.alignment = ALIGNMENT


  def generate_tonnage_formula(self, row, sets) -> None:
      #=SUM(PRODUCT(C34:C34),PRODUCT(C35:C35),PRODUCT(C36:C36)...)
      l = []

      first_row = row
      last_row = row + sets - 1
      for r in range(row, row + sets):
          l.append(f"{VOLUME_HEADERS['Load']['ColumnLetter']}{r}:{VOLUME_HEADERS['Reps']['ColumnLetter']}{r}")

      formula = '{}, {}{}), {})'.format(
          f"=IF(COUNT({VOLUME_HEADERS['Load']['ColumnLetter']}{first_row}:{VOLUME_HEADERS['Load']['ColumnLetter']}{last_row})>0",
          f"SUM(", "".join('PRODUCT({}), '.format(i) for i in l).rstrip(','),
          f"\"...\""
      )
      return formula


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
