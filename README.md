# Excel VBA UserForm - Data Entry

A simple Excel VBA UserForm for collecting and storing user information.

## Features

- **Clean Input Interface** — Name, Age, and City text fields
- **Add Button** — Submits data to Excel sheet automatically
- **Reset Button** — Clears all fields for next entry
- **Auto-Logging** — Entries saved to Sheet1 with timestamp
- **Module Trigger** — Launch UserForm via VBA module with F5

## UserForm Layout

| Control | Type | Purpose |
|---------|------|---------|
| NAME | TextBox | Enter person's name |
| AGE | TextBox | Enter age |
| CITY | TextBox | Enter city name |
| ADD | CommandButton | Save data to sheet |
| RESET | CommandButton | Clear all fields |

## How to Use

### Method 1: VBA Editor (Recommended)
1. Open `UserForm.xlsm` in Excel
2. Enable macros when prompted
3. Press `Alt + F11` to open **Visual Basic Editor**
4. Find the module in Project Explorer
5. Press **F5** to trigger and open the UserForm

### Method 2: Macro Dialog
1. Press `Alt + F8`
2. Select `ShowUserForm` macro
3. Click **Run**

### Method 3: Developer Tab
1. Go to **Developer → Macros**
2. Select `ShowUserForm`
3. Click **Run**

## Data Entry
1. Fill in **Name**, **Age**, **City**
2. Click **ADD** to save to sheet
3. Click **RESET** to clear fields for next entry

## Data Output

Entries are saved to Sheet1:

| Name | Age | City | Date |
|------|-----|------|------|
| Mark | 32 | LA | [auto] |
| Amy | 32 | LA | [auto] |

## Requirements

- Microsoft Excel 2010 or later
- Macros enabled
- Developer tab visible (File → Options → Customize Ribbon → Developer)

## VBA Structure
Project
├── Module1 (Launch Module)
│   └── Sub ShowUserForm() — Press F5 to run
└── UserForm4
├── NameTextBox (TextBox)
├── AgeTextBox (TextBox)
├── CityTextBox (TextBox)
├── AddButton (CommandButton)
└── ResetButton (CommandButton)
