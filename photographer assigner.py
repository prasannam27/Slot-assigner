import pandas as pd
from datetime import datetime, timedelta

# -------- Helper Functions --------
def parse_time_12h(t):
    t = str(t).strip().upper().replace('.', '')
    if "AM" in t or "PM" in t:
        t = t.replace("AM", " AM").replace("PM", " PM")
    return datetime.strptime(t.strip(), "%I:%M %p")

def get_slots(start, end):
    slots = []
    while start + timedelta(minutes=30) <= end:
        slots.append((start, start + timedelta(minutes=30)))
        start += timedelta(minutes=30)
    return slots

def is_available(free_list, slot):
    s_slot, e_slot = slot
    return any(s_free <= s_slot and e_free >= e_slot for s_free, e_free in free_list)

def occupy_slot(free_list, slot):
    s_slot, e_slot = slot
    updated = []
    for s_free, e_free in free_list:
        if e_free <= s_slot or s_free >= e_slot:
            updated.append((s_free, e_free))
        else:
            if s_free < s_slot:
                updated.append((s_free, s_slot))
            if e_slot < e_free:
                updated.append((e_slot, e_free))
    return sorted(updated, key=lambda x: x[0])

# -------- Step 1: Read Files --------
events_df = pd.read_excel(r"C:\Users\ADMIN\OneDrive\Desktop\events.xlsx")
photogs_df = pd.read_excel(r"C:\Users\ADMIN\OneDrive\Desktop\photographers.xlsx")

# -------- Step 2: Photographer Data --------
photographers = {}

for _, row in photogs_df.iterrows():
    free_periods = []
    cleaned = str(row["Free_Slots"]).replace("AM", " AM").replace("PM", " PM")

    for t in cleaned.split(","):
        if "-" in t:
            s, e = [x.strip() for x in t.replace("–", "-").split("-")]
            free_periods.append((parse_time_12h(s), parse_time_12h(e)))

    photographers[row["Name"]] = {
        "free": sorted(free_periods, key=lambda x: x[0]),
        "max_events": int(row["Max_Events"]),
        "assigned_events": set()
    }

# -------- Step 3: Assignment Logic --------
assignments = []

for _, e in events_df.iterrows():
    event_name = e["Event_Name"]
    t_range = str(e["Assigned_Time"]).replace("–", "-")
    start_str, end_str = [x.strip() for x in t_range.split("-")]

    start_time = parse_time_12h(start_str)
    end_time = parse_time_12h(end_str)
    event_slots = get_slots(start_time, end_time)

    assigned_photogs = set()

    for slot in event_slots:
        if len(assigned_photogs) >= 2:
            break

        for pid, pdata in sorted(
            photographers.items(),
            key=lambda x: len(x[1]["assigned_events"])
        ):
            if pid in assigned_photogs:
                continue
            if len(pdata["assigned_events"]) >= pdata["max_events"]:
                continue
            if is_available(pdata["free"], slot):

                assignments.append({
                    "Event_Name": event_name,
                    "Slot": f"{slot[0].strftime('%I:%M %p')} - {slot[1].strftime('%I:%M %p')}",
                    "Photographer": pid
                })

                pdata["free"] = occupy_slot(pdata["free"], slot)
                pdata["assigned_events"].add(event_name)
                assigned_photogs.add(pid)
                break

    while len(assigned_photogs) < 2:
        assignments.append({
            "Event_Name": event_name,
            "Slot": "No suitable slot",
            "Photographer": "None available"
        })
        assigned_photogs.add(f"NA_{len(assigned_photogs)}")

# -------- Step 4: Output --------
output_path = r"C:\Users\ADMIN\OneDrive\Desktop\assignments_output.xlsx"
pd.DataFrame(assignments).to_excel(output_path, index=False)

print(f"✅ Photographer assignment complete! Output saved to: {output_path}")
