
#SingleInstance force
#include ..\src\MenuEx.ahk

/**
 * In this example, we create a class, `TreeViewContextMenu`, that can be used to create a new
 * context menu object. Each instance of `TreeViewContextMenu` will initially include
 * two items "Copy value" and "Update value".
 *
 * This example also demonstrates how to use the item availability handler. The item availability handler
 * is a function that enables and disables menu items as a function of what was underneath the mouse
 * cursor when the user right-clicked, or what was selected when the user activated the context menu.
 * In our example, we do not want the option "Update value" to be active when the user right-clicks
 * on a TreeView item representing a property with a value that is an object, and so our item
 * availability handler disables / enables that option accordingly.
 */

demo()

class TreeViewContextMenu extends MenuEx {
    ; To effectively define default items that all instances of `TreeViewContextMenu` have, we must
    ; create a static method "__New". The reason we do this is because we are going to define an
    ; array on the prototype, so that way all instances of `TreeViewContextMenu` have access to the
    ; same array.
    static __New() {
        ; It's best to delete the static "__New" method because inheritors don't need to repeat these
        ; actions since inheritors will already have access to the array.
        this.DeleteProp('__New')
        ; Add the array to the prototype. These are referenced by `MenuEx.Prototype.Initialize`,
        ; which you can also override with new logic, but in his case we use the built-in logic.
        this.Prototype.DefaultItems := [
            ; Only the "Name" and "Value" are required. Other properties are "Options", which are the
            ; options described by https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add, and "Tooltip",
            ; which controls the tooltip text that will be displayed. Details about "Tooltip" are
            ; available in the description of `MenuExItem.Prototype.SetTooltipHandler`.

            ; Since we have defined the functions as class instance methods, we can define "Value"
            ; with the name of the method.
            { Name: 'Copy value', Value: 'SelectCopyValue' }
          , { Name: 'Update value', Value: 'SelectUpdateValue' }
        ]
    }
    ; Define an item availability handler to enable / disable "Update value" depending on what item
    ; was right-clicked on or what item was selected when the context menu was activated.
    HandlerItemAvailability(Ctrl, IsRightClick, Item, *) {
        text := ctrl.GetText(Item)
        if RegExMatch(text, '^[^=]+= \{ Object \}$') {
            this.__Item.Get('Update value').Disable()
        } else {
            this.__Item.Get('Update value').Enable()
        }
    }
    SelectCopyValue(Name, ItemPos, MenuObj, GuiObj, Ctrl, Item) {
        ; This function contains logic that uses the available information to add the property
        ; value to the clipboard

        ; Get the item's text
        text := ctrl.GetText(Item)
        ; Split the text at the equal sign; left side if the path, right side is the value
        split := StrSplit(text, '=', '`s')
        value := split[2]
        ; Add the value to the clipboard
        A_Clipboard := value
        ; Return text to display in the tooltip
        return 'Copied: ' value

    }
    SelectUpdateValue(Name, ItemPos, MenuObj, GuiObj, Ctrl, Item) {
        ; This function contains logic that uses the available information to change the value of
        ; the object's property.

        ; Get the item's text
        text := ctrl.GetText(Item)
        ; Split the text at the equal sign; left side is the path, right side is the value
        split := StrSplit(text, '=', '`s')
        path := split[1]
        value := split[2]
        ; Get input from user
        response := InputBox('Input a new value for ``' path '``', 'MenuEx example', , Trim(value, '"'))
        ; If they don't cancel
        if response.Result == 'OK' {
            posDot := InStr(path, '.', , , -1)
            ; Get a reference to the object
            obj := GetObjectFromString(SubStr(path, 1, posDot - 1))
            ; Get the property name
            prop := SubStr(path, posDot + 1)
            if IsNumber(response.Value) {
                ; Update the property value
                obj.%prop% := Number(response.Value)
                ; Update the TreeView item's text
                ctrl.Modify(Item, , path ' = ' response.Value)
                ; Return the text to display in the tooltip
                return 'Value updated to ' response.Value '.'
            } else {
                ; Update the property value
                obj.%prop% := response.Value
                ; Update the TreeView item's text
                ctrl.Modify(Item, , path ' = "' response.Value '"')
                ; Return the text to display in the tooltip
                return 'Value updated to "' response.Value '".'
            }
        ; If they cancel
        } else {
            ; Return the text to display in the tooltip
            return 'Update cancelled.'
        }
    }
}

class demo {
    static Call() {
        g := this.g := Gui('+Resize')
        tv := this.tv := g.Add('TreeView', 'w400 r10 vTv')
        this.ExampleObj := {
            Prop1: {
                Prop1_1: 'Val1_1'
              , Prop1_2: 'Val1_2'
            }
          , Prop2: 'Val2'
          , Prop3: {
                Prop3_1: {
                    Prop3_1_1: 'Val3_1_1'
                  , Prop3_1_2: 'Val3_1_2'
                }
            }
        }
        ids := [0]
        RecursiveFunction(this.ExampleObj, 'demo.ExampleObj')
        this.Options := { ShowTooltips: true }
        this.ContextMenu := TreeViewContextMenu(Menu(), this.Options)
        tv.OnEvent('ContextMenu', this.ContextMenu)
        g.Show()

        RecursiveFunction(Obj, Path) {
            for prop, val in ObjOwnProps(Obj) {
                NewPath := Path '.' prop

                if IsObject(val) {
                    ids.Push(tv.Add(NewPath ' = { Object }', ids[-1]))
                    RecursiveFunction(Val, NewPath)
                } else {
                    tv.Add(NewPath ' = "' val '"', ids[-1])
                }
            }
            ids.Pop()
        }
    }
}

/*
    Github: https://github.com/Nich-Cebolla/AutoHotkey-GetObjectFromString
    Author: Nich-Cebolla
    Version: 1.0.0
    License: MIT
*/

/**
 * @description - Converts a string path to an object reference. The object at the input path must
 * exist in the current scope of the function call.
 * @param {String} Str - The object path.
 * @param {Object} [InitialObj] - If set, the object path will be parsed as a property / item of
 * this object.
 * @returns {Object} - The object reference.
 */
GetObjectFromString(Str, InitialObj?) {
    static Pattern := '(?<=\.)[\w_\d]+(?COnProp)|\[\s*\K-?\d+(?COnIndex)|\[\s*(?<quote>[`'"])(?<key>.*?)(?<!``)(?:````)*\g{quote}(?COnKey)'
    if IsSet(InitialObj) {
        NewObj := InitialObj
        Pos := 1
        if SubStr(Str, 1, 1) !== '.' {
            Str := '.' Str
        }
    } else {
        RegExMatch(Str, '^[\w\d_]+', &InitialSegment)
        Pos := InitialSegment.Pos + InitialSegment.Len
        NewObj := %InitialSegment[0]%
    }
    while RegExMatch(Str, Pattern, &Match, Pos)
        Pos := Match.Pos + Match.Len

    return NewObj

    OnProp(Match, *) {
        NewObj := NewObj.%Match[0]%
    }
    OnIndex(Match, *) {
        NewObj := NewObj[Number(Match[0])]
    }
    OnKey(Match, *) {
        NewObj := NewObj[Match['key']]
    }
}
