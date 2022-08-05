# saysettha0.9
Version 0.9 of Saysettha Compiler for Visual Basic for Application

**1. **
> ```die()```
stop the program

**2.**
> ```$variable [= $variable2 / string / number / calc() / concat() ]``` 
 declare variable. The value will be flexible
 
 **3.**
> ```concat(list of string or $variables splited by comma)```
Join objects to make a string. 

**Example**: concat("Tom ","and Jerry") >> Tom and jerry

**4.**
> ```calc(expression)```
Do calculation. 

**Example**: calc(1-1+1) >> 1
**Note**:
1. `+  -  *  / ^ and s are allowed` 
2. `you can also use $variables`

**5.**
> ```goto <line>```
 
Jump to any line in Saysettha. 
**Example**: goto 4
**Note**:
1. `only integer numbers are allowed`
