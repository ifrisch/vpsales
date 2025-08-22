# MANUAL TASK SCHEDULER SETUP GUIDE

## 🎯 Setting up Van Paper Automation in Windows Task Scheduler

### Step 1: Open Task Scheduler
1. Press **Windows Key + R**
2. Type: `taskschd.msc`
3. Press **Enter**

### Step 2: Create Morning Task (7:32 AM CST)
1. In Task Scheduler, click **"Create Basic Task..."** (right side panel)
2. **Name**: `Van Paper Morning Automation`
3. **Description**: `Processes Van Paper 7:30 AM report at 7:32 AM CST`
4. Click **Next**

5. **Trigger**: Select **"Daily"**
6. Click **Next**

7. **Start time**: Set to **7:32:00 AM**
8. **Recur every**: `1 days`
9. Click **Next**

10. **Action**: Select **"Start a program"**
11. Click **Next**

12. **Program/script**: Browse and select:
    ```
    c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\run_morning_automation.bat
    ```
13. Click **Next**, then **Finish**

### Step 3: Create Afternoon Task (2:02 PM CST)
1. Click **"Create Basic Task..."** again
2. **Name**: `Van Paper Afternoon Automation`
3. **Description**: `Processes Van Paper 2:00 PM report at 2:02 PM CST`
4. Click **Next**

5. **Trigger**: Select **"Daily"**
6. Click **Next**

7. **Start time**: Set to **2:02:00 PM** (14:02)
8. **Recur every**: `1 days`
9. Click **Next**

10. **Action**: Select **"Start a program"**
11. Click **Next**

12. **Program/script**: Browse and select:
    ```
    c:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard\run_afternoon_automation.bat
    ```
13. Click **Next**, then **Finish**

### Step 4: Configure Task Settings (Important!)
For both tasks, right-click the task and select **"Properties"**:

1. **General Tab**:
   - Check **"Run whether user is logged on or not"**
   - Check **"Run with highest privileges"**

2. **Settings Tab**:
   - Check **"Allow task to be run on demand"**
   - Check **"Run task as soon as possible after a scheduled start is missed"**
   - Check **"If the task fails, restart every: 5 minutes"**
   - Set **"Attempt to restart up to: 3 times"**

3. Click **OK** and enter your Windows password if prompted

## ✅ VERIFICATION
1. You should see both tasks in Task Scheduler
2. Right-click either task and select **"Run"** to test
3. Check that your live app updates: https://vpsales.streamlit.app/

## 🎉 ALL DONE!
Your automation will now run:
- **7:32 AM CST** every day (processes 7:30 AM Van Paper report)
- **2:02 PM CST** every day (processes 2:00 PM Van Paper report)

The live Streamlit app will automatically update 1-2 minutes after each report is processed!
