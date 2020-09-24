This is a custom message box creator.It provides the following 
functionalities that are not provided by the MsgBox function in VB.

1.Allows you to change the font of the message.
2.Allows you to change the font color of the message.
3.Allows you to set italic font for the message.
4.Allows you to set bold font for the message.
5.Allows you to change the caption of the buttons.
6.You can have as  many buttons as you want on the message box.
7.The Message box has a totally new border style.
8.You can also set the AutoUnloadTime so that the MessageBox will automatically terminate when the specified time had elapsed returning the first button value.
9.The MessageBoxEx function used to create the messagebox returns the number of button pressed starting from zero.
10.Use the AddButton method to add as many buttons as you like. (Max 255 i think this is more than enough.)

Example
-------
AddButton 0, "&OK", True
AddButton 1, "&Cancel", , True
MessageBoxEx("Hello Everybody", , vbBlue, , 12, True, True, Msg_Left, True)


If you have any questions or problems please mail me at
vijaymp@rediffmail.com

Enjoy and have a nice day.

