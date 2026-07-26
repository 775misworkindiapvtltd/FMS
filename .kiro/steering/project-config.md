# FMS Project Configuration

## Google Sheet (FMS FASHION)

This is the live Google Sheet used by the FMS app for Login, Permissions (MASTER/STEPS), and future data:

```
https://docs.google.com/spreadsheets/d/10Aqav9bM_XQ28TkfjoXOoj8jRHenMerIEcx8gYkwo48/edit?usp=drive_web&ouid=114024121405657848674
```

Sheet ID: `10Aqav9bM_XQ28TkfjoXOoj8jRHenMerIEcx8gYkwo48`

### Known Tabs (from screenshots shared by user)
- `MASTER` — Sheet URL/Name config rows + LOGIN/PASSWORD table (row 4 = headers, data from row 5). Col A = LOGIN, Col B = PASSWORD, Col C = Name.
- `STEPS` — Permission mapping. Col A = Header/Category (e.g. "Utility FMS", "Copy of AMC FMS"), Col B = Sub-item/page name, Col C = User name who has permission.
- `DROPDOWN` — (purpose not yet defined, referenced tab name only)

Use this URL by default whenever testing/wiring the FMS Login page's "Sheet Configuration" fields, unless the user provides a different sheet.
