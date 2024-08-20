# Employee Scheduling Algorithm

On behalf of all variable-schedule employees and in protest of most mid-level managers, I've built a tool to automate the task of creating weekly schedules while optimizing for overall fairness and individual preferences.

[Link to tool](https://brananharrison.github.io/EmployeeScheduler/)

1) Download the input file
2) Add each employee and their availability/preferences, positions, shift info, labor needs, time-off requests, and attendance.
3) Save as Input.xlsx and upload
4) Polished schedule will download in 1-2 minutes <br><br>


## How it works

Within the Input Excel document, we find our Employees sheet containing data about each employee. Most data points are fixed by the manager at the hire date, such as eligible shifts/positions and availability. However, **columns labeled 'Desired' are able to be edited by respective employees at any time**, giving employees some autonomy in how their week is scheduled. 

![Img1](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched1.png)

![Img2](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched2.png)

Most clients share the input file on Google Drive with all employees to let them change preferences as needed. <br><br>


The other 3 sheets contain data for the manager to fill out, most importantly, the **Position Matrix**. Notice each position name is spaced exactly 10 cells apart, with a similar pattern within each position. This allows Python to scrape the data into the algorithm, both for the number of positions and individual number of shifts for each position, giving the manager **full freedom in designing shifts and positions**.

![Img3](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched3.png)

<br>
Time-off requests and attendance logs are inputted as well, and contribute to the algorithm.

![Img4](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched4.png)

![Img5](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched5.png) <br><br>


## The Algorithm

Each employee will have a variety of variables associated with a given shift:

### Boolean variables:
**Position Eligibility**: employee is eligible to work that position <br>
**Shift Eligibility**: employee is eligible to work that shift <br>
**Available**: the shift is within the employee's availability <br>
**Prefers Position**: the employee prefers that position for that day of the week <br>
**Prefers Shift**: the employee prefers that shift for that day of the week <br>
**Prefers OFF**: the employee prefers not to be scheduled for that day of the week <br>
**Requested Off**: the assignment of that shift would interfere with a TO request
‍
### Numerical variables:
**Max Hours Difference (MaxH)**: the difference between an employee's maximum desired hours and their current scheduled hours calculated INCLUDING the given shift <br>
**Min Hours Difference (MinH)**: minimum desired hours - scheduled hours <br>
**Max Days Difference (MaxD)**: max days - scheduled days <br>
**Min Days Difference (MinD)**: min days - scheduled days <br>

### Categorical variables:
**Request type**: if an employee requests time off it will either be (Approved) or (If possible) <br><br>


### Initial Scheduling Draft

Since the algorithm respects shift and hours preference, I chose these 5 categories of variable combinations. The first category contains the ideal case, where an employee prefers the shift, while the fifth contains the least ideal case, where an employee requested off for that day. Part of the optimization of this algorithm lies in the choices of these categories:

1. Employee is eligible, prefers the shift/position, has not requested off, and MaxD, MaxH > 0
2. Employee is eligible, does not prefer off, has not requested off, and MaxD, MaxH >= 0
3. Employee is eligible, prefers off, has not requested off, and MaxD, MaxH >= 0
4. Employee is eligible, does not prefer the shift/position, has not requested off, MaxD > 0, and MaxH > -5
5. Employee is eligible, has requested off with request type (if possible), and MaxD, MaxH >= 0

The assignment of each shift is done by iterating through each position, shift, and day of the week in order. It will then iterate down the 5 categories of preference until at least one viable employee is found, giving us 3 possible cases:

- If no employees are found it will iterate to the next preference category
- If one employee is found they are immediately assigned the shift
- If multiple employees are found a tiebreaker calculation is performed


### Tiebreaker calculation
In the event where more than one employee is eligible for the same shift, the first consideration is the Attendance Score. This is dependent on the 5 attendance categories from the manager's input sheet: No call no show, Inadmissible call-out, Admissible call-out, Write-up behavior, and Late.‍

```
AttendanceScore = minimum(0, 1 - (
0.6 * Number of no call no shows
0.2 * Number of inadmissible call-outs
0.15 * Number of write-up behaviors
0.03 * Number of late
0.02 * Number of admissible call-outs
))
```








### Test
```python
def greet(name):
    return f"Hello, {name}!"

print(greet('Branan'))
