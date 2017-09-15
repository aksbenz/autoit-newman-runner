# autoit-newman-runner
Autoit based newman runner. Creates newman command and executes in cmd.

Auto-It based UI to generate a newman command.

Allows to:
1. Select collection file
2. Select folders in collection
3. Select environment file
4. Select reporters
5. Select html template
6. Select SSL files and other ssl params
7. Set report folder
8. Set pre-cmds to run. Such as set proxy etc.

Based on above parameters generates the newman command. 
Command can be executed async in cmd window.
