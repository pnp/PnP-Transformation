# Introduction #
This folder contains documentation on InfoPath limitations in Office 365 and describes remediation models and approaches.

# InfoPath remediation models and approaches #
In this section we'll explain how one can transform InfoPath forms to solutions using alternative approaches like ASP.Net MVC, Single Page Apps using Knockout.js,...

The approach we've taken is to create a reference InfoPath form that implements the common patterns we've seen while working with several large customers. For each of these patterns we first explain the InfoPath implementation, followed by the implementation of that pattern using the alternative approaches. The screenshot is showing the reference form which implements an employee registration system.

![](http://i.imgur.com/esc3rMP.png)

The table below lists all the common patterns, the InfoPath implementation and the equivalent implementation for the alternative approaches (if there's an `X` in the table). Click on the listed pattern to learn more.

Common pattern | SPA using Knockout | ASP.Net MVC | ASP.Net Forms
---------------|:------------------:|:-----------:|:-----------:
[Populating fields on form load - set user information](https://github.com/OfficeDev/PnP-Transformation/blob/dev/InfoPath/Guidance/Patterns/Populating%20fields%20on%20form%20load-set%20user%20information.md) | x | x | x 
[Populating fields on form load - read list information](https://github.com/OfficeDev/PnP-Transformation/blob/dev/InfoPath/Guidance/Patterns/Populating%20fields%20on%20form%20load-read%20list%20information.md) | x | x | x 
Populating fields on form load - read list data | x | x | x 
[Submit the form via code](https://github.com/OfficeDev/PnP-Transformation/blob/dev/InfoPath/Guidance/Patterns/Submit%20the%20form%20via%20code.md) | x | x | x 
Switching view after form submission | x | x | x 
Retrieving user data | x | x | x 
[Read data collection and set multiple controls](https://github.com/OfficeDev/PnP-Transformation/blob/dev/InfoPath/Guidance/Patterns/Read%20data%20collection%20and%20set%20multiple%20controls.md) | x | x | x 
[Cascading data load](https://github.com/OfficeDev/PnP-Transformation/blob/dev/InfoPath/Guidance/Patterns/Cascading%20data%20load.md) | x | x | x 
Load existing item in form | x | x | x 


# InfoPath limitations and best practices #


