import coch

df_full = coch.extract_and_organise()
df_cal = coch.calendar_format(df_full)
df_hours = coch.hours_per_week(coch.show_only_shifts_new(df_full))

# coch.export_csv(df_cal)  # creates a calendar readable csv for import into google cal etc.
# coch.export_xls(df_full)  # Exports all fields to xls format
# coch.export_xls(df_hours)  # Exports calculation of hors per week to xls

print(df_full)
print(df_cal)
print(df_hours)
print("\nAverage Hours per Week:", df_hours["Hours"].mean())
