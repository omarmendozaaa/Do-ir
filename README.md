# DOit
Using pandas, fuzzywuzzy, openpyxl, tkinter and excel to take automated assistance 
# Motivation 🌸
My mother is an assistant in a college, she has 11 sections, 357 students. The information was not organized and I decided to help her. 
# How does it work?
We must first understand the amount of data that is stored in the attendance.
We use pandas to access and manipulate data in a more friendly way.
We realized that it was non-serialized data, so many errors can occur at the time of manipulation and even filling, so I decided to use fuzzywuzzy, a package that helps to calculate the percentage of coincidence or relationship that exists between two strings. In this way we mitigate errors when writing names, surnames, even writing only one name.
To access and edit within an excel file we use openpyxl in this way we can move in an excel file through its sheets, columns and rows.
I implemented an interface that allows me to test the commitment and eureka.
Now we go from arriving 11 sections, 365 students a day in more than 3 or 4 hours of work, to 3 minutes. 🌸
