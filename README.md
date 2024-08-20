# Employee Scheduling Algorithm

On behalf of all variable-schedule employees and in protest of most mid-level managers, I've built a tool to automate the task of creating weekly schedules while optimizing for overall fairness and individual preferences.

[Link to tool](https://brananharrison.github.io/EmployeeScheduler/)

1) Download the input file
2) Add each employee and their availability/preferences, positions, shift info, labor needs, time-off requests, and attendance.
3) Save as Input.xlsx and upload
4) Polished schedule will download in 1-2 minutes

## How it works

Within the Input Excel document, we find our Employees sheet, containing data about each employee. Most data points are fixed by the manager at the hire date, such as eligible shifts/positions and availability. However, columns labeled 'Desired' are able to be edited by respective employees at any time, giving employees some autonomy in how their week is scheduled.



### Test
```python
def greet(name):
    return f"Hello, {name}!"

print(greet('Branan'))
