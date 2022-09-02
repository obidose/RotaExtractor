import coch

full_df = coch.extract_and_organise()
# print(full_df)


# Convert Time Delta to hours for export - should ideally be added to module somewhere more appropriate
full_df["Shift Length"] = full_df["Shift Length"] / np.timedelta64(1, 'h')

coch.export_csv(coch.calendar_format(full_df))
