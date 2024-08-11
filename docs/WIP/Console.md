# `Console.frm`

## Introduction

### What is the VBA Console?
The main idea behind the console is used to log errors, print success or failure, ask the user for input, decide process with yes/no questions and run procedures (macros) and extras


### Preparation
To Prepare the console you have to do the following:  
If you want to activate Intellisense then go to the variable `Intellisense_Active` and set it to `true`. It is dependant on `Microsoft Visual Studio Extensebility 5.3`, because it uses VBComponents.  
If you want to activate stdLambda function go to the variable `stdLambda_Active` and set it to `true`. It is dependant sancarns stdVBA class `stdLambda`  
https://github.com/sancarn/stdVBA/tree/master  

Run `Console.Show`  
    This will initialize the console  

Now the Console can be used in process  

### Workmodes

#### 1. Logging
You can run `PrintConsole` to write Text to the current line or  
`PrintEnter` to write Text to the current line and paste a new one.  
Both function can take in a color/colorarray for the text

#### 2. User Input
There are 2 main User-interactions-  
    One is a message, followed by predeclared answers like yes/no,maybe or anything the programmer would like to use  
        VBA will continue to run until you write any of the available answers, else it will print, that your value is not allowed  
        If the input is allowed a user defined message may be shown  
    use `CheckPredeclaredAnswer` for this  
    The second one is a message, where it will ask you for a value of the user.  
        This might be combined with further errorhandling like checking if the input is of the right datatype, a wrong input will be shown accordingly (not yet implemented)  
    use `GetUserInput` for this  

#### 3. Running macros
This tool is strictly defined:  
    Write a Variable name  
    Enter |; |, as the seperator of arguments  
    Up to 30 additional arguments are allowed (1st argument is ProcedureName, all following arguments are arguments for said procedure)  
    Keep in mind, that Application.Run cannot call procedures form classes and userforms.  

#### 4. Extras
As of now there are the following special commands  
    `Help`  
        This will print a text explaining the functionalities of the console  
    `Clear`  
        This will clear all text of the Console  
    `Multiline`  
        This will allow the user to write several lines of code.  
        The last character has to be "_", otherwise it assumes the input is finished and it will combine the lines of code to a single line and run it  
    `Info`  
        This will get an explanation of the last run line  
    `?*`  
        This will input all following text as a string to a lambda function (not yet implemented)  
    `VariableName==*`  
        This will add another element in the `ConsoleVariables()` variable under the VariableName with the value to the right of the "==".  

### Extra Information
The console works with special modes:
```vb
     Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        PreDeclaredAnswer = 2
        UserLog = 3
        MultilineMode = 4
    End Enum
```
Logging is the basic one, where the console only recieves information  
UserInputt is variable input of the user  
PreDeclaredAnswer is for predeclared answers (duh)  
UserLog is for running procedures and extras  
MultilineMode allows for multiple lines to input(see Extras) 


## Console Public´s explained

```vb



' Public Console Functions

    ' Shows defined Message on Console
    ' Runs in loop
    ' Checks for userinput
    ' once typed check ifs of defined datatype
    ' keeps running until an acceptable input is inserted
    Public Function GetUserInput(Message As Variant, Optional InputType As String = "VARIANT") As Variant


    ' Shows defined Message on Console
    ' Runs in loop
    ' Checks for userinput
    ' once typed check ifs of allowed value
    ' keeps running until an acceptable input is inserted
    ' print successmessage (optional)
    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant


    ' Returns a String of the workbook-folderpath with a "Recognizer"
    Public Function PrintStarter() As Variant

    ' Prints a message to the console with defined color (or in_Basic) and inserts a new line
    '    Color needs to be either empty, 1 color or as many or more colors than Len(Text)
    Public Sub PrintEnter(Text As Variant, Optional Color As Variant)

    'Prints a message to the console with defined color (or in_Basic) without inserting a new line
    '    Color needs to be either empty, 1 color or as many or more colors than Len(Text)
    Public Sub PrintConsole(Text As Variant, Optional Color As Variant)
```

### Color

MIGHT NOT WORK ACCURATELY AT THIS STAGE  

Coloring isnt optimised at all, the more text you write the longer updating will take, because after every keyup event it will loop through all words and characters in said line to update its color.

The following things are currently defined for coloring while writing:  

```vb
    Private in_Basic       As Long
    Private in_Procedure   As Long
    Private in_Operator    As Long 'smooooooth operatooooor
    Private in_Datatype    As Long
    Private in_Value       As Long
    Private in_String      As Long
    Private in_Statement   As Long
    Private in_Keyword     As Long
    Private in_Parantheses As Long
    Private in_Variable    As Long
```

### Intellisense

Intellisense comes with a Listbox, which is grey.  
MIGHT NOT WORK PERFECTLY YET.  

While Initializing the Console it will go through all VBProjects and all VBCcomponents and get every Public-declared function, sub, variable and property.  

Dots are important for the lookup, as there is an `AstractionDepth` from 0 to 2 (Projects, Components, Public´s).  
Every Dot will increase it to 2 and more dots will result in nothing.  
Lookup will loop through all AbstractionDepth´s from Dot-Count up and paste the value he found into the listbox.  
To select it you need to press RIGHT, select it with DOWN/UP and press RIGHT again. To cancel press LEFT.  
If a word is recognized as a defined AbstractionDepth it will color it in the value of `in_Variable`

### Examples

From a Module
```vb
    Sub PrintToConsole()
        Dim Color(11) As Variant

        Color(00) = 255
        Color(01) = 65535
        Color(02) = 255
        Color(03) = 65535
        Color(04) = 255
        Color(05) = 65535
        Color(06) = 255
        Color(07) = 65535
        Color(08) = 255
        Color(09) = 65535
        Color(10) = 255
        Color(11) = 65535
        
        Console.Show
        ' Beware that PrintEnter adds vbcrlf to your text
        Console.PrintEnter "Hello World!"
        Console.PrintEnter "Hello There!", Color(0)
        Console.PrintConsole "THIS IS TEXT", Color

    End Sub

    Sub Square()
        
        Dim x As Variant

        Console.Show
        x = Console.GetUserInput("Write a Number: ", "DOUBLE") ' As of now, "DOUBLE" wont do anything, thats why the next line is here
        If IsNumeric(x) = True Then x = CDbl(x)

        x = x * x
        Console.PrintEnter x

    End Sub

    Sub Answer()
        
        Dim Message As Variant
        Dim AllowedValues(2) As Variant
        Dim Answers(2) As Variant

        Dim Value As Variant

        Message = "Would you like to square this Value? "
        AllowedValues(0) = "yes"
        AllowedValues(1) = "no"
        AllowedValues(2) = "cubic"
        Answers(0) = x * x
        Answers(1) = x
        Answers(2) = x * x * x

        Console.Show
        x = Console.GetUserInput("Write a Number: ", "DOUBLE") ' As of now, "DOUBLE" wont do anything, thats why the next line is here
        If IsNumeric(x) = True Then x = CDbl(x)
        
        x = Console.CheckPredeclaredAnswer(Message, AllowedValues, Answers)

    End Sub

    
```