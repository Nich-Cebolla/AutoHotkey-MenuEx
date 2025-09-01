/*
    Github: https://github.com/Nich-Cebolla/AutoHotkey-MenuEx
    Author: Nich-Cebolla
    Version: 1.0.0
    License: MIT
*/

/**
 * `MenuEx` is a composotion of AHK's native `Menu` class. The purpose of `MenuEx` is to provide a
 * standardized system for creating, modifying, and using a menu. Using `MenuEx` will feel more natural
 * to those who prefer an object-oriented coding style. For each item added to the menu, an associated
 * {@link MenuExItem} is created and added to the collection. The `MenuExItem` instances can be
 * accessed by name from the `MenuEx` instance, and the `MenuExItem` instance's properties can be
 * modified to change the characteristics of the menu item.
 *
 * ## Extending MenuEx
 *
 * `MenuEx` was designed with object inheritance in mind. One benefit of using `MenuEx` over
 * using `Menu` directly is it makes it easy to share menu items between menus and between scripts.
 *
 * Inheriting from `MenuEx` involves 2-4 steps. See the example in file
 * "test\demo-TreeView-context-menu.ahk" for a working example of each of these steps.
 *
 * 1. Define default items.
 *
 * To define default items, your class should define a static method "__New" that adds a property
 * "DefaultItems" to the prototype. "DefaultItems" is an array of objects, each object with required
 * properties { Name, Value } and optional properties { Options, Tooltip }.
 * - Name: The name of the menu item. This is used across the `MenuEx` class and related classes. It
 *   is the name that is used to get a reference to the `MenuExItem` instance associated with the
 *   menu item, e.g. `MenuExObj.Get("ItemName")`. It is also the text that is displayed in the menu
 *   for that item.
 * - Value: "Value" can be defined with three types of values.
 *   - A `Menu` object, if the menu item is a submenu.
 *   - A `Func` or callable object that will be called when the user selects the item.
 *   - A string representing the name of a class instance method defined by your custom class which
 *     inherits from `MenuEx` (see the "test\demo-TreeView-context-menu.ahk" for an example).
 * - Options: Any options as described in {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
 * - Tooltip: A value as described by {@link MenuExItem.Prototype.SetTooltipHandler}
 *
 * 2. Define "Initialize".
 *
 * Define a method "Initialize" which calls `this.AddObjectList(this.DefaultItems)` and any other
 * initialization tasks required by your class.
 *
 * 3. (Optional) Define instance methods.
 *
 * When creating a class that represents a menu that will be reused across various windows / scripts,
 * it makes sense to define the menu item functions directly in the class as instance methods.
 *
 * 4. (Optional) Define an item availability handler.
 *
 * It is often appropriate to adjust the availability of one or more menu items depending on the
 * context in which a context menu is activated. The item availability handler is only used when the
 * menu is a context menu (more specifically, the item availability handler is only used when
 * {@link MenuEx.Prototype.SetEventHandler} is called with a value of `1` or `2`.
 *
 */
class MenuEx {
    static __New() {
        this.DeleteProp('__New')
        proto := this.Prototype
        proto.__SelectionHandler := proto.__ItemAvailabilityHandler := proto.Token := proto.__TooltipHandler := ''
    }
    __New(Options?) {
        options := MenuEx.Options(Options ?? unset)
        this.SetSelectionHandler(options.SelectionHandler || unset)
        this.SetTooltipHandler(options.TooltipHandler || unset, options.TooltipDefaultOptions || unset)
        this.SetEventHandler(options.WhichMethod)
        this.SetItemAvailabilityHandler(options.ItemAvailabilityHandler)
        this.ShowTooltips := options.ShowTooltips
        this.__Item := MenuExItemCollection()
        this.__Item.CaseSense := options.CaseSense
        this.__Item.Default := ''
        this.Menu := options.IsMenuBar ? MenuBar() : Menu()
        this.Constructor := Class()
        this.Constructor.Base := MenuExItem
        this.Constructor.Prototype := {
            MenuEx: this
          , __Class: MenuExItem.Prototype.__Class
        }
        ObjSetBase(this.Constructor.Prototype, MenuExItem.Prototype)
        ObjRelease(ObjPtr(this))
        if HasMethod(this, 'Initialize') {
            this.Initialize(options)
        }
    }
    Add(Name, CallbackOrSubmenu, Options?, Tooltip?) {
        this.Menu.Add(Name, this.__SelectionHandler, Options ?? unset)
        this.__Item.Set(Name, this.Constructor.Call(Name, CallbackOrSubmenu, Options ?? unset, Tooltip ?? unset))
        return this.__Item.Get(Name)
    }
    /**
     * "AddList" should be used only if the menu which originally was associated with the items no
     * longer exists. To copy items from one menu to another, use "AddObjectList" instead.
     */
    AddList(Items) {
        container := this.__Item
        proto := this.Constructor.Prototype
        m := this.Menu
        for item in items {
            ObjSetBase(item, proto)
            container.Set(item.__Name, item)
            m.Add(item.__Name, item.__Value, item.__Options || unset)
        }
    }
    AddObject(Obj) {
        return this.Add(
            Obj.Name
          , Obj.Value
          , HasProp(Obj, 'Options') ? (Obj.Options || unset) : unset
          , HasProp(Obj, 'Tooltip') ? (Obj.Tooltip || unset) : unset
        )
    }
    AddObjectList(Objs) {
        for obj in Objs {
            this.Menu.Add(obj.Name, this.__SelectionHandler, HasProp(Obj, 'Options') ? (Obj.Options || unset) : unset)
            this.__Item.Set(obj.Name, this.Constructor.Call(
                obj.Name
              , obj.Value
              , HasProp(Obj, 'Options') ? (Obj.Options || unset) : unset
              , HasProp(Obj, 'Tooltip') ? (Obj.Tooltip || unset) : unset
            ))
        }
    }
    Clear() => this.__Item.Clear()
    Clone() => this.__Item.Clone()
    Delete(Name) {
        this.Menu.Delete(Name)
        this.__Item.Delete(Name)
    }
    Get(Name) => this.__Item.Get(Name)
    Has(Name) => this.__Item.Has(Name)
    Insert(InsertBefore, Name, CallbackOrSubmenu, Options?) {
        this.Menu.Insert(InsertBefore, Name, CallbackOrSubmenu, Options ?? unset)
        this.__Item.Set(Name, this.Constructor.Call(Name, CallbackOrSubmenu, Options ?? unset))
        return this.__Item.Get(Name)
    }
    OnSelect(Name, ItemPos, MenuObj) {
        if item := this.__Item.Get(Name) {
            params := { Menu: MenuObj, Name: Name, Pos: ItemPos, Token: this.Token }
            if IsObject(item.__Value) {
                result := item.__Value.Call(this, params)
            } else {
                result := this.%item.__Value%(params)
            }
            if this.ShowTooltips {
                if IsObject(item.Tooltip) {
                    str := item.Tooltip.Call(this, result)
                    if !IsObject(str) && StrLen(str) {
                        this.TooltipHandler.Call(str)
                    }
                } else if item.Tooltip {
                    this.TooltipHandler.Call(item.Tooltip)
                } else if !IsObject(result) && StrLen(result) {
                    this.TooltipHandler.Call(result)
                }
            }
        } else {
            throw UnsetItemError('Item not found.', -1, Name)
        }
    }
    Set(params*) => this.__Item.Set(params*)
    /**
     * @param {Integer} [Which = 0] - One of the following:
     * - 0: Use `0` when the menu is a `MenuBar` or a submenu. Generally, if the menu is not intended
     *   to be activated as a context menu, then `0` is appropriate.
     * - 1: Use `1` when the menu is activated as a context menu and the event handler is set to a
     *   control (not the gui).
     * - 2: Use `2` when the menu is activated as a context menu and the event handler is set to a
     *   gui window (not a control).
     */
    SetEventHandler(Which := 0) {
        if Which {
            this.DefineProp('Call', this.__GetOwner('__Call' Which))
        } else if this.HasOwnProp('Call') {
            this.DeleteProp('Call')
        }
    }
    SetItemAvailabilityHandler(Callback?) {
        this.__ItemAvailabilityHandler := Callback ?? ''
    }
    SetSelectionHandler(Callback?) {
        if this.HasOwnProp('__SelectionHandler')
        && this.__SelectionHandler.HasOwnProp('Name')
        && this.__SelectionHandler.Name == this.OnSelect.Name ' (bound)' {
            if IsSet(Callback) {
                ObjPtrAddRef(this)
                this.__SelectionHandler := Callback
            } else {
                OutputDebug('The current selection handler is already set to ``' this.__SelectionHandler.Name '``.`n')
            }
        } else if IsSet(Callback) {
            this.__SelectionHandler := Callback
        } else {
            ; This creates a reference cycle.
            this.__SelectionHandler := ObjBindMethod(this, 'OnSelect')
            ; This is to identify that the object is the bound method (and thus requires handling
            ; the reference cycle).
            this.__SelectionHandler.DefineProp('Name', { Value: this.OnSelect.Name ' (bound)' })
        }
    }
    SetTooltipHandler(Callback?, DefaultOptions?) {
        this.__TooltipHandler := Callback ?? MenuEx.TooltipHandler(DefaultOptions ?? unset)
    }
    __Call1(GuiCtrlObj, Item, IsRightClick, X, Y) {
        this.Token := {
            Ctrl: GuiCtrlObj, Gui: GuiCtrlObj.Gui
          , IsRightClick: IsRightClick
          , Item: Item, X: X, Y: Y
        }
        ObjSetBase(this.Token, MenuExActivationToken.Prototype)
        if HasMethod(this, 'ItemAvailabilityHandler') {
            this.ItemAvailabilityHandler()
        } else if IsObject(this.__ItemAvailabilityHandler) {
            this.__ItemAvailabilityHandler.Call(this)
        }
        this.Menu.Show(X, Y)
    }
    __Call2(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
        this.Token := {
            Ctrl: GuiCtrlObj, Gui: GuiObj
          , IsRightClick: IsRightClick
          , Item: Item, X: X, Y: Y
        }
        ObjSetBase(this.Token, MenuExActivationToken.Prototype)
        if HasMethod(this, 'ItemAvailabilityHandler') {
            this.ItemAvailabilityHandler()
        } else if IsObject(this.__ItemAvailabilityHandler) {
            this.__ItemAvailabilityHandler.Call(this)
        }
        this.Menu.Show(X, Y)
    }
    __Delete() {
        if this.HasOwnProp('Constructor')
        && this.Constructor.HasOwnProp('Prototype')
        && this.Constructor.Prototype.HasOwnProp('MenuEx') {
            ObjPtrAddRef(this)
            this.DeleteProp('Constructor')
        }
        if this.HasOwnProp('__SelectionHandler')
        && this.__SelectionHandler.HasOwnProp('Name')
        && this.__SelectionHandler.Name == this.OnSelect.Name ' (bound)' {
            ObjPtrAddRef(this)
            this.DeleteProp('__SelectionHandler')
        }
    }
    __Enum(VarCount) => this.__Item.__Enum(VarCount)
    __GetOwner(Prop, ReturnDesc := true) {
        b := this
        while b {
            if b.HasOwnProp(Prop) {
                break
            }
            b := b.Base
        }
        if !b {
            throw PropertyError('Property not found.', -1, Prop)
        }
        return ReturnDesc ? b.GetOwnPropDesc(Prop) : b
    }
    Capacity {
        Get => this.__Item.Capacity
        Set => this.__Item.Capacity := Value
    }
    CaseSense => this.__Item.CaseSense
    Count => this.__Item.Count
    IsMenuBar => this.Menu is MenuBar
    ItemAvailabilityHandler {
        Get => this.__ItemAvailabilityHandler
        Set => this.SetItemAvailabilityHandler(Value)
    }
    Handle => this.Menu.Handle
    SelectionHandler {
        Get => this.__SelectionHandler
        Set => this.SetSelectionHandler(Value)
    }
    TooltipHandler {
        Get => this.__TooltipHandler
        Set => this.SetTooltipHandler(Value)
    }

    class Options {
        static Default := {
            CaseSense: false
          , TooltipHandler: ''
          , SelectionHandler: ''
          , ShowTooltips: false
          , WhichMethod: 1
          , TooltipDefaultOptions: ''
          , ItemAvailabilityHandler: ''
          , IsMenuBar: false
        }
        static Call(Options?) {
            if IsSet(Options) {
                o := {}
                d := this.Default
                for prop in d.OwnProps() {
                    o.%prop% := HasProp(Options, prop) ? Options.%prop% : d.%prop%
                }
                return o
            } else {
                return this.Default.Clone()
            }
        }
    }

    class TooltipHandler {
        /**
         * By default, `MenuEx.TooltipHandler.Numbers` is an array with integers 1-20, and is used to track which
         * tooltip id numbers are available and which are in use. If tooltips are created from multiple
         * sources, then the list is invalid because it may not know about every existing tooltip. To
         * overcome this, `MenuEx.TooltipHandler.Numbers` can be set with an array that is shared by other objects,
         * sharing the pool of available id numbers.
         *
         * All instances of `MenuEx.TooltipHandler` will inherently draw from the same array, and so calling
         * `MenuEx.TooltipHandler.SetNumbersList` is unnecessary if the objects handling tooltip creation are all
         * `MenuEx.TooltipHandler` objects.
         */
        static SetNumbersList(List) {
            this.Numbers := List
        }
        static DefaultOptions := {
            Duration: 2000
          , X: 0
          , Y: 0
          , Mode: 'Mouse' ; Mouse / Absolute (M/A)
        }
        static Numbers := [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]

        /**
         * @param {Object} [DefaultOptions] - An object with zero or more options as property : value pairs.
         * These options are used when a corresponding option is not passed to {@link MenuEx.TooltipHandler.Prototype.Call}.
         * @param {Integer} [DefaultOptions.Duration = 2000] - The duration in milliseconds for which the
         * tooltip displayed. A value of 0 causes the tooltip to b e dislpayed indefinitely until
         * {@link MenuEx.TooltipHandler.Prototype.End} is called with the tooltip number. Negative and positive
         * values are treated the same.
         * @param {Integer} [DefaultOptions.X = 0] - If `DefaultOptions.Mode == "Mouse"` (or "M"), a number
         * of pixels to add to the X-coordinate. If `DefaultOptions.Mode == "Absolute"` (or "A"), the
         * X-coordinate relative to the screen.
         * @param {Integer} [DefaultOptions.Y = 0] - If `DefaultOptions.Mode == "Mouse"` (or "M"), a number
         * of pixels to add to the Y-coordinate. If `DefaultOptions.Mode == "Absolute"` (or "A"), the
         * Y-coordinate relative to the screen.
         * @param {String} [DefaultOptions.Mode = "Mouse"] - One of the following:
         * - "Mouse" or "M" - The tooltip is displayed near the mouse cursor.
         * - "Absolute" or "A" - The tooltip is displayed at the screen coordinates indicated by the
         * options.
         */
        __New(DefaultOptions?) {
            if IsSet(DefaultOptions) {
                o := this.DefaultOptions := {}
                d := MenuEx.TooltipHandler.DefaultOptions
                for p in d.OwnProps()  {
                    o.%p% := HasProp(DefaultOptions, p) ? DefaultOptions.%p% : d.%p%
                }
            } else {
                this.DefaultOptions := MenuEx.TooltipHandler.DefaultOptions.Clone()
            }
        }
        /**
         * @param {String} Str - The string to display.
         *
         * @param {Object} [Options] - An object with zero or more options as property : value pairs.
         * If a value is absent, the corresponding value from `TooltipHandlerObj.DefaultOptions` is used.
         * @param {Integer} [Options.Duration] - The duration in milliseconds for which the
         * tooltip displayed. A value of 0 causes the tooltip to b e dislpayed indefinitely until
         * {@link MenuEx.TooltipHandler.Prototype.End} is called with the tooltip number. Negative and positive
         * values are treated the same.
         * @param {Integer} [Options.X] - If `Options.Mode == "Mouse"` (or "M"), a number
         * of pixels to add to the X-coordinate. If `Options.Mode == "Absolute"` (or "A"), the
         * X-coordinate relative to the screen.
         * @param {Integer} [Options.Y] - If `Options.Mode == "Mouse"` (or "M"), a number
         * of piYels to add to the Y-coordinate. If `Options.Mode == "Absolute"` (or "A"), the
         * Y-coordinate relative to the screen.
         * @param {String} [Options.Mode] - One of the following:
         * - "Mouse" or "M" - The tooltip is displayed near the mouse cursor.
         * - "Absolute" or "A" - The tooltip is displayed at the screen coordinates indicated by the
         * options.
         *
         * @returns {Integer} - The tooltip number used for the tooltip. If the duration is 0, pass
         * the number to {@link MenuEx.TooltipHandler.Prototype.End} to end the tooltip. Otherwise, you do not need
         * to save the tooltip number, but the tooltip number can be used to target the tooltip when calling
         * `ToolTip`.
         */
        Call(Str, Options?) {
            if MenuEx.TooltipHandler.Numbers.Length {
                n := MenuEx.TooltipHandler.Numbers.Pop()
            } else {
                throw Error('The maximum number of concurrent tooltips has been reached.', -1)
            }
            if IsSet(Options) {
                Get := _Get1
            } else {
                Get := _Get2
            }
            T := CoordMode('Tooltip', 'Screen')
            switch SubStr(Get('Mode'), 1, 1), 0 {
                case 'M':
                    M := CoordMode('Mouse', 'Screen')
                    MouseGetPos(&X, &Y)
                    ToolTip(Str, X + Get('X'), Y + Get('Y'), n)
                    SetTimer(ObjBindMethod(this, 'End', n), -Abs(Get('Duration')))
                    CoordMode('Mouse', M)
                case 'A':
                    ToolTip(Str, Get('X'), Get('Y'), n)
                    SetTimer(ObjBindMethod(this, 'End', n), -Abs(Get('Duration')))
            }
            CoordMode('Tooltip', T)

            return n

            _Get1(prop) {
                return HasProp(Options, prop) ? Options.%prop% : this.DefaultOptions.%prop%
            }
            _Get2(prop) {
                return this.DefaultOptions.%prop%
            }
        }
        End(n) {
            ToolTip(,,,n)
            MenuEx.TooltipHandler.Numbers.Push(n)
        }
        /**
         * @param {Object} [DefaultOptions] - An object with zero or more options as property : value pairs.
         * These options are used when a corresponding option is not passed to {@link MenuEx.TooltipHandler.Prototype.Call}.
         * The existing default options are overwritten with the new object.
         */
        SetDefaultOptions(DefaultOptions) {
            this.DefaultOptions := DefaultOptions
        }
    }
}

class MenuExItem {
    static __New() {
        this.DeleteProp('__New')
        proto := this.Prototype
        proto.__Name := proto.__Value := proto.__Options := proto.Tooltip := ''
    }
    /**
     * @see {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {String} Name - The text to display on the menu item. Although the AutoHotkey
     * documents indicate that, for `Menu.Prototype.Add`, the paramter `MenuItemName` can also be
     * the position of an existing item, that is not applicable here; only pass the name to this
     * parameter.
     * @param {*} CallbackOrSubmenu - The function to call as a new thread when the menu item is
     * selected, or a reference to a Menu object to use as a submenu.
     *
     * Regarding functions:
     *
     * The function can be any callable object. When using this library (`MenuEx` and related classes),
     * the functions are not called directly when the user selects a menu item; a handler function is
     * called which then access the `MenuExItem` object associated with the menu item that was selected.
     * The function is then called from the property "__Value".
     *
     * If `CallbackOrSubmenu` is a function, then the function should accept two parameters:
     * 1. The {@link MenuEx} instance.
     * 2. An object with properties:
     *   - Menu: The menu object.
     *   - Name: The menu item name that was selected.
     *   - Pos: The position of the menu item that was selected.
     *   - Token: The {@link MenuExActivationToken} instance that was created when the menu was
     *     activated. "Token" only has a significant value when {@link MenuEx.Prototype.SetEventHandler}
     *     was called with a value of `1` or `2`. That is, if the menu is not activated as a context
     *     menu, then "Token" is always an empty string. If the menu is activated as a context menu,
     *     then "Token" is a {@link MenuExActivationToken} instance.
     *
     * The function can also be the name of a class method. For details about this, see the section
     * "Extending MenuEx" in the description above {@link MenuEx}.
     *
     * The function's return value may be used if {@link MenuEx#ShowTooltips} is nonzero. For details
     * about how the return value is used, see {@link MenuExItem.Prototype.SetTooltipHandler}.
     *
     * @param {String} [MenuItemOptions = ""] - Any options as described in
     * {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {*} [Tooltip] - See {@link MenuExItem.Prototype.SetTooltipHandler} for details about
     * this parameter.
     */
    __New(Name, CallbackOrSubmenu, MenuItemOptions := '', Tooltip?) {
        this.__Name := Name
        this.__Value := CallbackOrSubmenu
        this.__Options := MenuItemOptions
        if IsSet(Tooltip) {
            this.Tooltip := Tooltip
        }
    }
    Check() {
        this.MenuEx.Menu.Check(this.__Name)
    }
    Delete() {
        this.MenuEx.Menu.Delete(this.__Name)
        this.MenuEx.__Item.Delete(this.__Name)
    }
    Disable() {
        this.MenuEx.Menu.Disable(this.__Name)
    }
    Enable() {
        this.MenuEx.Menu.Enable(this.__Name)
    }
    SetIcon(FileName, IconNumber := 1, IconWidth?) {
        this.MenuEx.Menu.SetIcon(this.__Name, FileName, IconNumber, IconWidth ?? unset)
    }
    /**
     * When {@link MenuEx#ShowTooltips} is true, there are three approaches for controlling the tooltip
     * text that is displayed when the user selects a menu item. When the user selects a menu item,
     * the return value returned by the function associated with the menu item is stored in a variable,
     * and the property {@link MenuExItem#Tooltip} is evaluated to determine if a tooltip will be displayed,
     * and if so, what the text will be.
     *
     * Note that, by default, the value of {@link MenuExItem#Tooltip} is an empty string.
     *
     * If {@link MenuExItem#Tooltip} is an object, it is assumed to be a function or callable object.
     * The function is called with parameters:
     * 1. The {@link MenuEx} instance.
     * 2. The return value from the menu item's function.
     *
     * The function should return the string that will be displayed by the tooltip. If the function
     * returns an object or an empty string, no tooltip is displayed.
     *
     * If {@link MenuExItem#Tooltip} is a significant string value, the return value from the menu
     * item's function is ignored and {@link MenuExItem#Tooltip} is displayed in the tooltip.
     *
     * If {@link MenuExItem#Tooltip} is zero or an empty string, and if the return value from the menu
     * item's function is a number or non-empty string, the return value is displayed in the tooltip.
     * Note that if the return value is a numeric zero or a string containing only a zero, that is
     * displayed in the tooltip; only an empty string will cause a tooltip to not be displayed.
     *
     * @param {*} Value - A value to one of the effects described by the description.
     */
    SetTooltipHandler(Value) {
        /**
         * See {@link MenuExItem.Prototype.SetTooltipHandler} for details.
         * @memberof MenuExItem
         * @instance
         */
        this.Tooltip := Value
    }
    ToggleCheck() {
        this.MenuEx.Menu.ToggleCheck(this.__Name)
    }
    ToggleEnable() {
        this.MenuEx.Menu.ToggleEnable(this.__Name)
    }
    Uncheck() {
        this.MenuEx.Menu.Uncheck(this.__Name)
    }
    Name {
        Get => this.__Name
        Set {
            this.MenuEx.Menu.Rename(this.__Name, Value)
            this.MenuEx.__Item.Delete(this.__Name)
            this.__Name := Value
            this.MenuEx.__Item.Set(Value, this)
        }
    }
    Options {
        Get => this.__Options
        Set {
            this.MenuEx.Menu.Add(this.__Name, , Value)
            this.__Options := Value
        }
    }
    Value {
        Get => this.__Value
        Set {
            this.__Value := Value
        }
    }
}


class MenuExActivationToken {
    __New(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
        this.Gui := GuiObj
        this.Ctrl := GuiCtrlObj
        this.Item := Item
        this.IsRightClick := IsRightClick
        this.X := X
        this.Y := Y
    }
}

class MenuExItemCollection extends Map {
}
