# Excel VBA Employee Management System

This project is a simple **Employee Management System built in Excel using VBA**. It allows you to **add, and delete employee records** efficiently, directly in Excel, without manually formatting or managing tables.

---

## **Project Features**

1. **Add Unique Employees**
   - Adds employee records from a source sheet (`Sheet1`) to a styled `OutputSheet`.
   - Avoids duplicate EmployeeIDs.
   - Highlights the **first column** (headers) in **dark magenta with white bold text**.
   - Data column has **dark magenta font**.
   - Each employee record is displayed as a **vertical table** with an **outer border in dark magenta**.
   - **Messages displayed:**
     - `"Missing values"` → if any field is empty.
     - `"Duplicate record"` → if EmployeeID already exists.
     - `"Record(s) updated successfully!"` → after all valid records are added.

2. **Delete Employee Record**
   - Deletes an employee record from `OutputSheet` by **EmployeeID**.
   - EmployeeID to delete is entered in **cell H11** on `Sheet1`.
   - Confirms deletion with a **“Are you sure you want to delete?”** message.
   - Deletes the **entire table block** and shifts remaining tables up.
   - Clears **H11** automatically after deletion.
   - Shows confirmation message: `"Record deleted"` or `"EmployeeID not found"`.

3. **Data Styling & Formatting**
   - All employee tables have **outer borders** in dark magenta.
   - Header column is bold, white text on dark magenta background.
   - Data column is dark magenta text.
   - Proper alignment and spacing for readability.

---

## **Step-by-Step Implementation**

### **Step 1: Prepare the Source Sheet**
1. Create `Sheet1` for input data.
2. Add column headers in Row 1:
3. Enter employee data starting from Row 2.

---

### **Step 2: Create the Output Sheet**
1. Create `OutputSheet`.
2. This sheet will automatically display all employee records added via the macro.

---

### **Step 3: Add the AddUniqueEmployees Macro**
1. Open VBA Editor (`Alt + F11`), insert a **module**, and paste the macro.
2. The macro will:
- Loop through all employees in `Sheet1`.
- Skip duplicates or missing data.
- Add new tables in `OutputSheet` with proper styling and borders.
- Show a single “Record(s) updated successfully!” message after completion.

---

### **Step 4: Add the DeleteEmployeeRecord Macro**
1. Paste the macro in the same module.
2. Workflow:
- Enter EmployeeID in **H11** on `Sheet1`.
- Click the **Delete button** assigned to this macro.
- Confirm deletion in the prompt.
- Macro deletes the table and clears H11 automatically.

---

### **Step 5: Assign Buttons (Optional)**
1. Go to `Developer → Insert → Button`.
2. Assign **AddUniqueEmployees** macro to an “Add Record” button.
3. Assign **DeleteEmployeeRecord** macro to a “Delete Record” button.

---

### **Step 6: Test the System**
1. Add a few employees in `Sheet1`.
2. Click the **Add Record** button.
3. Verify that the table is generated correctly in `OutputSheet`.
4. Enter an EmployeeID in H11, click **Delete Record**, and confirm deletion.

---

## **Future Enhancements**
- Add **Update Record** functionality.
- Add **search functionality** by EmployeeID or Name.
- Export tables to **PDF or another workbook**.
- Create a **mini Excel dashboard** for employee statistics.

---

## **Screenshots (Example)**
| Input Sheet | OutputSheet |
|------------|-------------|
| ![Input Example](example_input.png) | ![Output Example](example_output.png) |

---

## **Conclusion**
This project demonstrates how **Excel VBA can be used to manage employee data efficiently**, combining **automation, styling, and user-friendly prompts**. It’s a great example for anyone learning **Excel macros and data automation**.
