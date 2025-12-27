# Slot-assigner
A Python-based scheduling system that assigns photographers to events based on time availability, workload limits, and fairness. The tool uses 30-minute time slots, prevents duplicate assignments, and outputs results to an Excel file for easy coordination.

Features
- Assigns 2 photographers per event
- Uses 30-minute time slot segmentation
- Respects individual availability windows
- Enforces maximum events per photographer
- Prevents duplicate photographer assignments for the same event
- Fair workload distribution (least-loaded photographer first)
- Excel-based input and output for easy integration

How It Works
1. Reads event timings and photographer availability from Excel files
2. Splits event durations into 30-minute slots
3. Checks photographer availability for each slot
4. Assigns photographers while enforcing constraints
5. Outputs final assignments to an Excel file

Tech Stack
- Python 3
- Pandas
- Datetime
- Excel (openpyxl)
