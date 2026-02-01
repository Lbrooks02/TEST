import sys
import winsound
import inspect
from typing import get_origin
from datetime import datetime

def Validate_Data(*functionParameters: str) -> bool:    #Validate data format of required parameters
    """
        Parameters:
            :functionParameters: *args containing str parameter names of calling function
             Supports module-level functions and class methods (instance, class, and static). Does not support nested/locally-defined functions
        Return: Bool indicating process failure/success
        Usage: Supports module-level functions and class methods (instance, class, and static). Does not support nested/locally-defined functions
        Ex:
            def userFunction(name: str, age: int):
                Validate_Data("name", "age")
    """
    PRIMITIVE_DATATYPES = (str, int, float, bool)
    CONTAINER_DATATYPES = (list, dict, tuple, set)
    VALID_DATATYPES = PRIMITIVE_DATATYPES + CONTAINER_DATATYPES + (datetime,)
    soundPlayed: bool = False                           #Ensures sound plays only once per function call

    def _Play_Sound(soundType: str):                    #Plays sound based on process success/failure (helper sub-function)
        if not soundType:
            sys.exit(">> _Play_Sound 'soundType' must be non-empty")
        elif not isinstance(soundType, str):
            sys.exit(">> _Play_Sound 'soundType' must be of type str")
        if soundType.lower() == "success":
            soundSource = r"C:\Windows\Media\Alarm03.wav"
        elif soundType.lower() == "error":
            soundSource = r"C:\Windows\Media\chord.wav"
        else:
            return None
        winsound.PlaySound(soundSource, winsound.SND_FILENAME | winsound.SND_NODEFAULT)

    def _Error_Handler(errorMessage: str = ">> Error - (Validate_Data) - Unknown", forceQuit: bool = True, playSound: bool = True):  #Prints error message + quits/returns success status (helper sub-function)
        nonlocal soundPlayed #nonlocal keyword allows nested function to access variable from immediate outer function, instead of making new local variable
        if (not soundPlayed) and (playSound):
            _Play_Sound("Error")
            soundPlayed = True
        if forceQuit:
            sys.exit(errorMessage)
        print(errorMessage)
        return False
    
    for i, parameter in enumerate(functionParameters, 1):       #Check *functionParameters are correctly inserted into Validate_Data()
        if (not isinstance(parameter, str)) or (not parameter.strip()):
            _Error_Handler(f">> Error - (Validate_Data): Argument {i} '{parameter}' must be a non-empty str")
    callerFunctionFrame = inspect.currentframe().f_back         #Get frame object of calling function (verifies this function is nested 1 level within another function)
    if not callerFunctionFrame:                     
        _Error_Handler(">> Error - (Validate_Data): Unable to determine caller function")  
        return False
    callerFunctionName = callerFunctionFrame.f_code.co_name     #Get name of caller function
    parameterDataDict: dict = callerFunctionFrame.f_locals      #Maps parameter variables of calling function to the corresponding argument values
    callerFunctionObject = None
    if "self" in parameterDataDict:                                 #Check if Validate_Data is called from instance method
        callerObject = parameterDataDict["self"]                    #Stores instance object that called the instance method
        rawCallerFunctionObject = type(callerObject).__dict__.get(callerFunctionName)   #Stores the calling instance method defined in the class
        if rawCallerFunctionObject is None:
            rawCallerFunctionObject = getattr(type(callerObject), callerFunctionName, None)         #Check parent class for inherited methods
    elif "cls" in parameterDataDict:                                #Check if Validate_Data is called from class method              
        callerObject = parameterDataDict["cls"]                     #Stores class object
        rawCallerFunctionObject = callerObject.__dict__.get(callerFunctionName)         #Stores the calling class method defined in the class
        if rawCallerFunctionObject is None:
            rawCallerFunctionObject = getattr(callerObject, callerFunctionName, None)   #Check parent class for inherited methods
    elif callerFunctionFrame.f_globals.get(callerFunctionName):     #Handling if Validate_Data is called from function
        rawCallerFunctionObject = callerFunctionFrame.f_globals.get(callerFunctionName)  
    else:                                                           #Handling if Validate_Data is called from static method
        qualifiedName = callerFunctionFrame.f_code.co_qualname      #Returns className.methodName 
        if "." in qualifiedName:
            className = qualifiedName.split(".")[0]                 #Extract the class name
            classObject = callerFunctionFrame.f_globals.get(className)  #Retrieve class object
            if classObject:
                rawCallerFunctionObject = classObject.__dict__.get(callerFunctionName)  #Retrieve function name from class object
    if not rawCallerFunctionObject:
        _Error_Handler(f">> Error - ({callerFunctionName}): Unable to determine caller function object")
    if (isinstance(rawCallerFunctionObject, classmethod)) or (isinstance(rawCallerFunctionObject, staticmethod)):
        callerFunctionObject = rawCallerFunctionObject.__func__                         #Access rawCallerFunctionObject function of the method
    else:
        callerFunctionObject = rawCallerFunctionObject
    if callerFunctionObject:                                    #Get the calling function's object
        parameterAnnotations = getattr(callerFunctionObject, "__annotations__", {}) #Stores {parameterName: typeHint} data from the calling function
        functionDetails = inspect.signature(callerFunctionObject)   #Get the signature object of the calling function (container that holds parameter default values)
        parameterDefaults = {}
        for name, param in functionDetails.parameters.items():
            if param.default is not inspect.Parameter.empty:    #Track {parameter name: default value} data from the calling function
                parameterDefaults[name] = param.default
    else:
        parameterAnnotations = {}
        parameterDefaults = {}
    if not functionParameters:
        _Error_Handler(f">> Error - ({callerFunctionName}): No parameter names provided for validation")
    if not parameterAnnotations:                                #Warn user if calling function has no type hints for *functionParameters
        _Error_Handler(f">> Warning - ({callerFunctionName}): No parameters in this function have type annotations", forceQuit=False, playSound=False)
    for parameterName in functionParameters:
        if parameterName not in parameterDataDict:              #Confirm all elements in *functionParameters exist as parameters in calling function
            _Error_Handler(f">> Error - ({callerFunctionName}): '{parameterName}' parameter not found in calling function")
            continue
        parameterValue = parameterDataDict[parameterName]
        if (parameterName in parameterDefaults) and (parameterValue == parameterDefaults[parameterName]): #Skip *functionParameters with default values
            _Error_Handler(f">> Warning - {callerFunctionName}: '{parameterName}' parameter contains a default value and will not be validated", forceQuit=False, playSound=False)
            continue
        if parameterValue in (None, "", [], {}, ()):            #Catch any *functionParameters with blank/invalid values
            _Error_Handler(f">> Error - ({callerFunctionName}): '{parameterName}' argument must be non-empty")
            continue
        if (parameterAnnotations) and (parameterName not in parameterAnnotations): #All elements in *functionParameters must have an associated type hint, or none should. Mixing both in one Validate_Data call is not supported
            _Error_Handler(f">> Warning - ({callerFunctionName}): '{parameterName}' parameter is missing a type annotation", forceQuit=False, playSound=False)
            continue
        if parameterName in parameterAnnotations:
            requiredDataType = parameterAnnotations[parameterName]
            if get_origin(requiredDataType):                    #Tracks data type for more advanced types (Ex. list[int] errors with isinstance, but returns "list type/class" using get_origin)
                requiredDataType = get_origin(requiredDataType)
            strDataTypeRequired: str = getattr(requiredDataType, "__name__", str(requiredDataType)) #Convert data type/class object to string
            if requiredDataType in VALID_DATATYPES:             #Confirm value for *functionParameters element matches type hint
                if requiredDataType in PRIMITIVE_DATATYPES:     #Logic for primitive data types
                    if type(parameterValue) is not requiredDataType:            #Compare parameter's actual type with data type from type hint 
                        strDataTypeActual: str = type(parameterValue).__name__  #Convert data type/class object to string
                        _Error_Handler(f">> Error - ({callerFunctionName}): '{parameterName}' argument must be of type {strDataTypeRequired}, got {strDataTypeActual}") 
                else:                                           #Logic for container data types and datetime
                    if not isinstance(parameterValue, requiredDataType):    
                        strDataTypeActual: str = type(parameterValue).__name__
                        _Error_Handler(f">> Error - ({callerFunctionName}): '{parameterName}' argument must be of type {strDataTypeRequired}, got {strDataTypeActual}") 
            else:                                               #Catch any attempts to use unsupported data types
                _Error_Handler(f">> Error - ({callerFunctionName}): '{parameterName}' parameter uses unsupported type {strDataTypeRequired}") 
    return True