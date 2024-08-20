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
- If multiple employees are found a tiebreaker calculation is performed <br><br>

### Tiebreaker calculation
In the likely event where more than one employee is eligible for the same shift, we must consider 3 seemingly intangible qualities: **attendance, seniority, and balance of shift assignments**. <br><br>

#### The first consideration is the **Attendance Score**
This is dependent on the 5 attendance categories from the manager's input sheet: No call no show, Inadmissible call-out, Admissible call-out, Write-up behavior, and Late.‍

```
AttendanceScore = minimum(0, 1 - (
    0.6 * Number of no call no shows
    0.2 * Number of inadmissible call-outs
    0.15 * Number of write-up behaviors
    0.03 * Number of late
    0.02 * Number of admissible call-outs
))
```
The result is a robust measurement of an employee's attendance and behavior on the interval [0,1] where 1 is a perfect employee, and 0 is a much less than perfect employee. <br><br>

#### In addition to the attendance score, the second consideration is the **Assigned Hours Proportion**
This is a proportion created from the current assigned hours / minimum desired hours and it's meant to balance the distribution of shift assignments based on what has already been assigned.

```python
for employee in employees:
    if employee['Name'] in preferred_list:
        if employee['Min Desired Hours'] != 0:
            if employee['Hours assigned'] / employee['Min Desired Hours'] > largest_proportion:
                largest_proportion = employee['Hours assigned'] / employee['Min Desired Hours']
for employee in employees:
    if employee['Name'] in preferred_list:
        if employee['Min Desired Hours'] == 0:
            comparison_score = 0.01
        else:
            comparison_score = (employee['Hours assigned'] / employee[
                'Min Desired Hours']) / largest_proportion
        comparison_list.append((employee['Name'], comparison_score))
winner = min_score(comparison_list)
```

Notice the comparison score is calculated not just by **(hours assigned) / (min desired hours)**, but by taking that quotient against the largest proportion of the employees in the list. This gives us a special distribution on the continuous interval [0,1], where 1 is always the employee with the largest hours proportion, and proportions less than that are each with respect to that maximum, **effectively normalizing each proportion to a distance from 1 on the interval (0,1]**. <br><br>

#### The third and final component is **Seniority**
Seniority is calculated simply by an employee's duration of hire / the most senior employee in the comparison pool. Similar to the assigned hours proportion, seniority values lie along the interval (0,1]. <br><br>

Now that we have 3 measures, **attendance, assigned hours proportion,** and **seniority**, we can use these to choose the ideal assignment in the event of a tiebreaker calculation. <br><br>

## Probabilistic Assignments
Since seniority and attendance are metrics calculated over a large time window, they are unlikely to change much over time. To avoid schedules feelign stagnant or certain employees being stuck with undesireable shifts, I've employed a probabilistic approach to keep things interesting. To keep things simple for now, let's refer to an employee's "comparison score" to be any of the 3 scores we defined above (attendance, assigned hours proportion, and seniority).

Let {X<sub>1</sub>, X<sub>2</sub>} be 2 employees' comparison scores on the interval [0,1]. Without loss of generality, suppose X<sub>1</sub> > X<sub>2</sub>. Define Z = 5*(X<sub>1</sub> - X<sub>2</sub> ) as the difference between the scores with a coefficient of 5 in preparation for a composition with the sigmoid function. A higher coefficient would result in more skew towards the leader, while a lower would do the opposite, however I think 5 is a fair trade-off. Finally we utilize the function below to decide the probability of the leader being chosen. A greater difference in comparison score will yield a higher probability of the leader being chosen.

![Img6](https://github.com/brananharrison/EmployeeScheduler/blob/master/img/sched6.png)


