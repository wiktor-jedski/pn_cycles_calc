# pn_cycles_calc
Flight hours/cycles/days controls calculator for parts installed on A/C.
Script is based on Excel reports taken from CAMO software and returns a report with total hours, cycles and days for
specific controls for a given component.

## Input/output files formatting
A/C report column names:
AC | FLIGHT_DAT | OFF_HOUR | ON_HOUR | FLIGHT_HOURS | FLIGHT_MINS | CYCLES | TOTAL_AC_FLIGHT_HOUR | TOTAL_AC_FLIGHT_MIN
| TOTAL_AC_CYCLES | FLIGHT_LOG

P/N report column names:
PN | SN | CONTROL | SCHEDULE_HOURS | SCHEDULE_CYCLES | SCHEDULE_DAYS | ACTUAL_HOURS | ACTUAL_CYCLES | ACTUAL_DAYS
| INSTALLED_ON | INSTALLED_DATE

Output report column names:
PN | SN | CONTROL | SCHEDULE_HOURS | SCHEDULE_CYCLES | SCHEDULE_DAYS | ACTUAL_HOURS | ACTUAL_CYCLES | ACTUAL_DAYS
| INSTALLED_ON | INSTALLED_DATE | TOTAL_HRS | TOTAL_CYC | TOTAL_DAYS

## Requirements
openpyxl==3.0.10
