#!/usr/bin/env python3
# Author: Jon Schipp <jonschipp@gmail.com, jschipp@illinois.edu>
import argparse
from Workout import Workout
from Utils import Utils


def usage():
  doc = '''
  Could not open file! Does it exist? Is it valid JSON?

  A file containing the API credentials in JSON should be read in using ``-f <file>''
  Its contents should be formatted like this:

  {
    "user":"aristauser",
    "password":"asdfasdfasdfasdf"
  }
  '''[1:]
  return doc


def arguments():
  parser = argparse.ArgumentParser(description='A generator for customizable workout templates using spreadsheets.')
  parser.add_argument("-W", "--weeks",     type=int, help="Number of weeks in the program (def: 8)")
  parser.add_argument("-F", "--frequency", type=int, help="Training frequency in number of days per week (def: 3)")
  parser.add_argument("-S", "--slots",     type=int, help="Number of exercises slots per workout (def: 3)")
  parser.add_argument("-s", "--sets",     type=int, help="Number of sets per slot (def: 10)")
  parser.add_argument("-f", "--filename", type=str, help="Spreadsheet output filename, (def: workout.xlsx)")
  args = parser.parse_args()

  weeks = args.weeks
  frequency = args.frequency
  slots = args.slots
  sets = args.sets
  filename = args.filename

  return(weeks, frequency, slots, sets, filename)


def main():
  weeks, frequency, slots, sets, filename = arguments()

  Program = Workout()
  Program.generate_weeks(weeks=weeks)
  Program.generate_frequency(frequency=frequency)
  Program.generate_slots(slots=slots, sets=sets, frequency=frequency)
  Utils.clear(workbook=Program.wb)
  Utils.save(workbook=Program.wb, filename=filename)


if __name__ == "__main__":
  main()
