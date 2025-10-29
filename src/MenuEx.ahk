/*
    Github: https://github.com/Nich-Cebolla/AutoHotkey-MenuEx
    Author: Nich-Cebolla
    License: MIT
*/

/**
 * `MenuEx` is a composotion of AHK's native `Menu` class. The purpose of `MenuEx` is to provide a
 * standardized system for creating, modifying, and using a menu. For each item added to the menu,
 * an associated {@link MenuExItem} is created and added to the collection. The `MenuExItem` instances
 * can be accessed by name from the `MenuEx` instance, and the `MenuExItem` instance's properties can
 * be modified to change the characteristics of the menu item.
 *
 * ## Context Menu
 *
 * Though `MenuEx` is useful for any menu, I designed it with a focus on functionality related to
 * context menus. When creating a context menu with `MenuEx`, the `MenuEx` instance will have a
 * method "Call" which activates the context menu. To use, simply pass the `MenuEx` object to the
 * event handler for the gui or control.
 *
 * @example
 *  g := Gui()
 *  MenuExObj := MenuEx(Menu())
 *  g.OnEvent('ContextMenu', MenuExObj) ; pass `MenuExObj` to event handler
 * @
 *
 * Or
 *
 * @example
 *  g := Gui()
 *  g.Add('TreeView', 'w100 r10 vTv')
 *  MenuExObj := MenuEx(Menu())
 *  g['Tv'].OnEvent('ContextMenu', MenuExObj) ; pass `MenuExObj` to event handler
 * @
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
 * {@link MenuEx.Prototype.SetEventHandler} is called with a value of `1` or `2`).
 *
 * Define the item availability handler as an instance method "HandlerItemAvailability".
 *
 */
class MenuEx {
    static __New() {
        this.DeleteProp('__New')
        MenuEx_SetConstants()
        proto := this.Prototype
        proto.__HandlerSelection := proto.__HandlerItemAvailability := proto.Token :=
        proto.__HandlerTooltip := ''
    }
    /**
     * {@link https://learn.microsoft.com/en-us/windows/win32/menurc/wm-changeuistate}
     *
     * @param {Integer} Handle - The handle of the window to receive the message. This would typically
     * be the parent window so all child windows can have the same UI state, but in the context of
     * an AHK menu `Hwnd` can also be the menu's handle, i.e. `MenuObj.Handle`.
     *
     * @param {Integer} Action - One of the following:
     * - UIS_CLAS - The UI state flags specified by the high-order word should be cleared.
     * - UIS_INITIALIZE - The UI state flags specified by the high-order word should be changed
     *   based on the last input event. For more information, see Remarks.
     * - UIS_SET - The UI state flags specified by the high-order word should be set.
     *
     * @param {Integer} State - One or more of the following. Combine with bitwise "or" ( | ),
     * e.g. `State := UISF_ACTIVE | UISF_HIDEACCEL`.
     * - UISF_ACTIVE - A control should be drawn in the style used for active controls.
     * - UISF_HIDEACCEL - Keyboard accelerators are hidden.
     * - UISF_HIDEFOCUS - Focus indicators are hidden.
     */
    static ChangeUiState(Handle, Action, State) {
        SendMessage(WM_CHANGEUISTATE, (State & 0xFFFF) << 16 | (Action & 0xFFFF), 0, , Handle)
    }
    /**
     * {@link https://learn.microsoft.com/en-us/windows/win32/menurc/wm-queryuistate}
     *
     * @param {Integer} Handle - The handle of the window to receive the message.
     *
     * @param {VarRef} [OutActive] - A variable that will receive either:
     * - 0 if the UISF_ACTIVE flag was not included in the return value.
     * - 1 if the UISF_ACTIVE flag was included in the return value.
     *
     * @param {VarRef} [OutAccel] - A variable that will receive either:
     * - 0 if the UISF_HIDEACCEL flag was not included in the return value. This would indicate
     *   that keyboard accelerators are visible.
     * - 1 if the UISF_HIDEACCEL flag was included in the return value. This would indicate that
     *   keyboard accelerators are hidden.
     *
     * @param {VarRef} [OutFocus] - A variable that will receive either:
     * - 0 if the UISF_HIDEFOCUS flag was not included in the return value.  This would indicate
     *   that focus indicators are visible.
     * - 1 if the UISF_HIDEFOCUS flag was included in the return value. This would indicate that
     *   focus indicators are hidden.
     */
    static QueryUiState(Handle, &OutActive?, &OutAccel?, &OutFocus?) {
        if value := SendMessage(WM_QUERYUISTATE, 0, 0, , Handle) {
            OutActive := value & UISF_ACTIVE
            OutAccel := value & UISF_HIDEACCEL
            OutFocus := value & UISF_HIDEFOCUS
        } else {
            OutActive := OutAccel := OutFocus := 0
        }
    }
    /**
     * {@link https://learn.microsoft.com/en-us/windows/win32/menurc/wm-updateuistate}
     *
     * @param {Integer} Handle - The handle of the window to receive the message. This would typically
     * be the parent window so all child windows can have the same UI state, but in the context of
     * an AHK menu `Hwnd` can also be the menu's handle, i.e. `MenuObj.Handle`.
     *
     * @param {Integer} Action - One of the following:
     * - UIS_CLAS - The UI state flags specified by the high-order word should be cleared.
     * - UIS_INITIALIZE - The UI state flags specified by the high-order word should be changed
     *   based on the last input event. For more information, see Remarks.
     * - UIS_SET - The UI state flags specified by the high-order word should be set.
     *
     * @param {Integer} State - One or more of the following. Combine with bitwise "or" ( | ),
     * e.g. `State := UISF_ACTIVE | UISF_HIDEACCEL`.
     * - UISF_ACTIVE - A control should be drawn in the style used for active controls.
     * - UISF_HIDEACCEL - Keyboard accelerators are hidden.
     * - UISF_HIDEFOCUS - Focus indicators are hidden.
     */
    static UpdateUiState(Handle, Action, State) {
        SendMessage(WM_UPDATEUISTATE, (State & 0xFFFF) << 16 | (Action & 0xFFFF), 0, , Handle)
    }
    /**
     * @param {Menu|MenuBar} [MenuObj] - The menu object. If unset, a new instance of `Menu` is created.
     *
     * @param {Object} [Options] - An object with zero or more options as property : value pairs.
     *
     * @param {Boolean} [Options.CaseSense = false] - If true, the collection is case-sensitive. This
     * means that accessing menu items from the collection by name is case-sensitive.
     *
     * @param {*} [Options.HandlerTooltip = ""] - See {@link MenuEx.Prototype.SetTooltipHandler~Callback}.
     *
     * @param {*} [Options.HandlerSelection = ""] - See {@link MenuEx.Prototype.SetSelectionHandler~Callback}.
     *
     * @param {Boolean} [Options.ShowTooltips = false] - If true, enables tooltip functionality.
     * `MenuEx`'s tooltip functionality allows you to define your menu and related options to
     * display a tooltip when the user selects a menu item. See {@link MenuExItem.Prototype.SetTooltipHandler}
     * for details and see file "test\demo-TreeView-context-menu.ahk" for a working example.
     *
     * @param {Integer} [Options.WhichMethod = 1] - `Options.WhichMethod` is passed directly to
     * method {@link MenuEx.Prototype.SetEventHandler}. See the description for details.
     *
     * @param {Object} [Options.TooltipDefaultOptions = ""] - The value passed to the second parameter
     * of {@link MenuEx.TooltipHandler} when creating the tooltip handler function object. If
     * `Options.HandlerTooltip` is set with a function, then `Options.TooltipDefaultOptions` is
     * ignored.
     *
     * @param {*} [Options.HandlerItemAvailability = ""] - See
     * {@link MenuEx.Prototype.SetItemAvailabilityHandler~Callback}.
     */
    __New(MenuObj?, Options?) {
        this.Menu := MenuObj ?? Menu()
        options := MenuEx.Options(Options ?? unset)
        this.SetSelectionHandler(options.HandlerSelection || unset)
        this.SetTooltipHandler(options.HandlerTooltip || unset, options.TooltipDefaultOptions || unset)
        this.SetEventHandler(options.WhichMethod)
        this.SetItemAvailabilityHandler(options.HandlerItemAvailability)
        this.ShowTooltips := options.ShowTooltips
        this.__Item := MenuExItemCollection()
        this.__Item.CaseSense := options.CaseSense
        this.__Item.Default := ''
        this.Constructor := Class()
        this.Constructor.Base := MenuExItem
        this.Constructor.Prototype := {
            MenuEx: this
          , __Class: MenuExItem.Prototype.__Class
        }
        ObjSetBase(this.Constructor.Prototype, MenuExItem.Prototype)
        ObjRelease(ObjPtr(this))
        this.Initialize(options)
    }
    /**
     * @param {String} Name - The name of the menu item. This is used across the {@link MenuEx} class
     * and related classes. It is the name that is used to get a reference to the {@link MenuExItem}
     * instance associated with the menu item, e.g. `MenuExObj.Get("ItemName")`. It is also the text
     * that is displayed in the menu for that item. It is also the value assigned to the "__Name"
     * property of the {@link MenuExItem} instance.
     *
     * @param {*} CallbackOrSubmenu - One of the following:
     * - A `Menu` object, if the menu item is a submenu.
     * - A `Func` or callable object that will be called when the user selects the item.
     * - A string representing the name of a class instance method defined by your custom class which
     *   inherits from `MenuEx` (see the "test\demo-TreeView-context-menu.ahk" for an example).
     *
     * The value of `CallbackOrSubmenu` is assigned to the "__Value" property of the {@link MenuExItem}
     * instance.
     *
     * @param {String} [Options] - The options as described in
     * {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {*} [HandlerTooltip] - The tooltip handler options as described in
     * {@link MenuExItem.Prototype.SetTooltipHandler}.
     *
     * @returns {MenuExItem}
     */
    Add(Name, CallbackOrSubmenu, Options?, HandlerTooltip?) {
        this.Menu.Add(Name, this.__HandlerSelection, Options ?? unset)
        this.__Item.Set(Name, this.Constructor.Call(Name, CallbackOrSubmenu, Options ?? unset, HandlerTooltip ?? unset))
        return this.__Item.Get(Name)
    }
    /**
     * "AddList" should be used only if the menu which originally was associated with the items no
     * longer exists. To copy items from one menu to another, use "AddObjectList" instead.
     * @param {MenuExItem[]} Items - An array of {@link MenuExItem} objects. For each item in the
     * array, the base of the item is changed to {@link MenuEx#Constructor.Prototype} and the item
     * is added to the menu.
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
    /**
     * @param {Object} Obj - An object with parameters as property : value pairs.
     * - Name: The name of the menu item. This is the value passed to the first parameter "Name" of
     *   {@link MenuEx.Prototype.Add} and is the value that is set to property "__Name"
     *   of the {@link MenuExItem} instance.
     * - Value: The value of the menu item; this is the value passed to the second parameter
     *   "CallbackOrSubmenu" of {@link MenuEx.Prototype.Add} and is the value that is set to property
     *   "__Value" of the {@link MenuExItem} instance.
     * - Options: The options for the menu item. This is the value passed to the third parameter
     *   "Options" of {@link MenuEx.Prototype.Add} and is the value that is set to property
     *   "__Options" of the {@link MenuExItem} instance.
     * - Tooltip: The tooltip options for the menu item. This is the value passed to the fourth parameter
     *   "HandlerTooltip" of {@link MenuEx.Prototype.Add} and is the value that is set to property
     *   "__HandlerTooltip" of the {@link MenuExItem} instance.
     * @returns {MenuExItem}
     */
    AddObject(Obj) {
        return this.Add(
            Obj.Name
          , Obj.Value
          , HasProp(Obj, 'Options') ? (Obj.Options || unset) : unset
          , HasProp(Obj, 'Tooltip') ? (Obj.Tooltip || unset) : unset
        )
    }
    /**
     * @param {Object[]|MenuExItem[]} Objs - An array of objects as described by {@link MenuEx.Prototype.AddObject},
     * or an array of {@link MenuExItem} instance objects.
     */
    AddObjectList(Objs) {
        m := this.Menu
        items := this.__Item
        constructor := this.Constructor
        handlerSelection := this.__HandlerSelection
        for obj in Objs {
            m.Add(obj.Name, handlerSelection, HasProp(Obj, 'Options') ? (Obj.Options || '') : unset)
            items.Set(obj.Name, constructor(
                obj.Name
              , obj.Value
              , HasProp(Obj, 'Options') ? (Obj.Options || unset) : unset
              , HasProp(Obj, 'Tooltip') ? (Obj.Tooltip || unset) : unset
            ))
        }
    }
    Delete(Name) {
        this.Menu.Delete(Name)
        this.__Item.Delete(Name)
    }
    DeleteList(Names) {
        m := this.Menu
        items := this.__Item
        for name in Names {
            m.Delete(name)
            items.Delete(name)
        }
    }
    DeletePattern(NamePattern) {
        m := this.Menu
        items := this.__Item
        names := []
        for name in items {
            if RegExMatch(name, NamePattern) {
                names.Push(name)
            }
        }
        for name in names {
            m.Delete(name)
            items.Delete(name)
        }
    }
    Get(Name) => this.__Item.Get(Name)
    Has(Name) => this.__Item.Has(Name)
    Initialize(*) {
        if HasProp(this, 'DefaultItems') {
            this.AddObjectList(this.DefaultItems)
        }
    }
    /**
     * @param {String|Integer} InsertBefore - The name or position of the menu item before which to
     * insert the new menu item.
     *
     * @param {String} Name - The name of the menu item. This is used across the {@link MenuEx} class
     * and related classes. It is the name that is used to get a reference to the {@link MenuExItem}
     * instance associated with the menu item, e.g. `MenuExObj.Get("ItemName")`. It is also the text
     * that is displayed in the menu for that item.
     *
     * @param {*} CallbackOrSubmenu - One of the following:
     * - A `Menu` object, if the menu item is a submenu.
     * - A `Func` or callable object that will be called when the user selects the item.
     * - A string representing the name of a class instance method defined by your custom class which
     *   inherits from `MenuEx` (see the "test\demo-TreeView-context-menu.ahk" for an example).
     *
     * @param {String} [Options] - The options as described in
     * {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {*} [HandlerTooltip] - The tooltip handler options as described in
     * {@link MenuExItem.Prototype.SetTooltipHandler}.
     *
     * @returns {MenuExItem}
     */
    Insert(InsertBefore, Name, CallbackOrSubmenu, Options?, HandlerTooltip?) {
        this.Menu.Insert(InsertBefore, Name, CallbackOrSubmenu, Options ?? unset)
        this.__Item.Set(Name, this.Constructor.Call(Name, CallbackOrSubmenu, Options ?? unset, HandlerTooltip ?? unset))
        return this.__Item.Get(Name)
    }
    /**
     * "OnSelect" is the default selection handler that is called when the user selects a menu item.
     * Your code will not call "OnSelect" directly.
     * @param {String} Name - The name of the menu item that was selected.
     * @param {Integer} ItemPos - The item position of the menu item that was selected.
     * @param {Menu} MenuObj - The menu object associated wit hthe menu item that was selected.
     */
    OnSelect(Name, ItemPos, MenuObj) {
        if item := this.__Item.Get(Name) {
            if token := this.Token {
                this.Token := ''
                if IsObject(item.__Value) {
                    result := item.__Value.Call(this, Name, ItemPos, MenuObj, token.Gui, token.Ctrl, token.Item)
                } else {
                    result := this.%item.__Value%(Name, ItemPos, MenuObj, token.Gui, token.Ctrl, token.Item)
                }
            } else {
                if IsObject(item.__Value) {
                    result := item.__Value.Call(this, Name, ItemPos, MenuObj)
                } else {
                    result := this.%item.__Value%(Name, ItemPos, MenuObj)
                }
            }
            if this.ShowTooltips {
                if IsObject(item.__HandlerTooltip) {
                    str := item.__HandlerTooltip.Call(this, result)
                    if !IsObject(str) && StrLen(str) {
                        this.__HandlerTooltip.Call(str)
                    }
                } else if item.__HandlerTooltip {
                    this.__HandlerTooltip.Call(item.__HandlerTooltip)
                } else if !IsObject(result) && StrLen(result) {
                    this.__HandlerTooltip.Call(result)
                }
            }
        } else {
            throw UnsetItemError('Item not found.', -1, Name)
        }
    }
    /**
     * See {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#SetColor}.
     */
    SetColor(ColorValue, ApplyToSubmenus := true) {
        this.Menu.SetColor(ColorValue, ApplyToSubmenus)
    }
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
    /**
     * @param {*} [Callback] - A `Func` or callable object that is called prior to showing the menu,
     * intended to enable or disable menu items depending on the item that was underneath the cursor
     * when the use right-clicked, or the item that was selected when the user activated the
     * context menu. The item availability handler is only called if
     * {@link MenuEx.Prototype.SetEventHandler} was called with a value of `1` or `2`. If `Callback`
     * is unset, the value of property "__ItemAvailabilityHandle" is set with an empty string, which
     * causes the process to not call an item availability handler.
     */
    SetItemAvailabilityHandler(Callback?) {
        this.__HandlerItemAvailability := Callback ?? ''
    }
    /**
     * @param {*} [Callback] - A `Func` or callable object that is called when the user selects a
     * menu item. The `Callback` is unset, the selection handler is defined as the "OnSelect" method,
     * which should be suitable for most use cases.
     */
    SetSelectionHandler(Callback?) {
        if this.HasOwnProp('__HandlerSelection')
        && this.__HandlerSelection.HasOwnProp('Name')
        && this.__HandlerSelection.Name == this.OnSelect.Name ' (bound)' {
            if IsSet(Callback) {
                ObjPtrAddRef(this)
                this.__HandlerSelection := Callback
            } else {
                OutputDebug('The current selection handler is already set to ``' this.__HandlerSelection.Name '``.`n')
            }
        } else if IsSet(Callback) {
            this.__HandlerSelection := Callback
        } else {
            ; This creates a reference cycle.
            this.__HandlerSelection := ObjBindMethod(this, 'OnSelect')
            ObjRelease(ObjPtr(this))
            ; This is to identify that the object is the bound method (and thus requires handling
            ; the reference cycle).
            this.__HandlerSelection.DefineProp('Name', { Value: this.OnSelect.Name ' (bound)' })
        }
    }
    /**
     * @param {*} [Callback] - A `Func` or callable object that is called after the function associated
     * with a menu item returns. `Callback` is expected to display a tooltip informing the user of
     * the result of the action associated with the menu item the user selected. For details about
     * this process, see {@link MenuExItem.Prototype.SetTooltipHandler}. If `Callback` is unset,
     * the property "__HandlerTooltip" is set with an instance of
     * {@link MenuEx.TooltipHandler} which should be suitable for most use cases.
     * @param {Object} [DefaultOptions] - An object with property : value pairs representing the
     * options to pass to the {@link MenuEx.TooltipHandler} constructor.
     */
    SetTooltipHandler(Callback?, DefaultOptions?) {
        this.__HandlerTooltip := Callback ?? MenuEx.TooltipHandler(DefaultOptions ?? unset)
    }
    /**
     * See {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Show}.
     */
    Show(X?, Y?) {
        this.Menu.Show(X ?? unset, Y ?? unset)
    }
    __Call1(Ctrl, Item, IsRightClick, X, Y) {
        this.Token := {
            Ctrl: Ctrl
          , Item: Item
          , Gui: Ctrl.Gui
        }
        if HasMethod(this, 'HandlerItemAvailability') {
            this.HandlerItemAvailability(Ctrl, IsRightClick, Item, X, Y)
        } else if IsObject(this.__HandlerItemAvailability) {
            this.__HandlerItemAvailability.Call(this, Ctrl, IsRightClick, Item, X, Y)
        }
        this.Menu.Show(X, Y)
    }
    __Call2(GuiObj, Ctrl, Item, IsRightClick, X, Y) {
        this.Token := {
            Ctrl: Ctrl
          , Gui: GuiObj
          , Item: Item
        }
        if HasMethod(this, 'HandlerItemAvailability') {
            this.HandlerItemAvailability(GuiObj, Ctrl, IsRightClick, Item, X, Y)
        } else if IsObject(this.__HandlerItemAvailability) {
            this.__HandlerItemAvailability.Call(this, GuiObj, Ctrl, IsRightClick, Item, X, Y)
        }
        CoordMode('Menu', 'Screen')
        this.Menu.Show(X, Y)
    }
    __Delete() {
        if this.HasOwnProp('Constructor')
        && this.Constructor.HasOwnProp('Prototype')
        && this.Constructor.Prototype.HasOwnProp('MenuEx') {
            ObjPtrAddRef(this)
            this.DeleteProp('Constructor')
        }
        if this.HasOwnProp('__HandlerSelection')
        && this.__HandlerSelection.HasOwnProp('Name')
        && this.__HandlerSelection.Name == this.OnSelect.Name ' (bound)' {
            ObjPtrAddRef(this)
            this.DeleteProp('__HandlerSelection')
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
    HandlerItemAvailability {
        Get => this.__HandlerItemAvailability
        Set => this.SetItemAvailabilityHandler(Value)
    }
    Handle => this.Menu.Handle
    HandlerSelection {
        Get => this.__HandlerSelection
        Set => this.SetSelectionHandler(Value)
    }
    HandlerTooltip {
        Get => this.__HandlerTooltip
        Set => this.SetTooltipHandler(Value)
    }

    class Options {
        static Default := {
            CaseSense: false
          , HandlerItemAvailability: ''
          , HandlerSelection: ''
          , HandlerTooltip: ''
          , ShowTooltips: false
          , TooltipDefaultOptions: ''
          , WhichMethod: 1
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
            Duration: 3000
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
        proto.__Name := proto.__Value := proto.__Options := proto.__HandlerTooltip := ''
    }
    /**
     * @see {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {String} Name - The text to display on the menu item. Although the AutoHotkey
     * documents indicate that, for `Menu.Prototype.Add`, the paramter `MenuItemName` can also be
     * the position of an existing item, that is not applicable here; only pass the name to this
     * parameter.
     *
     * @param {*} CallbackOrSubmenu - The function to call as a new thread when the menu item is
     * selected, or a reference to a Menu object to use as a submenu.
     *
     * Regarding functions:
     *
     * The function can be any callable object.
     *
     * If `CallbackOrSubmenu` is a function, then the function can accept the following parameters:
     * 1. {MenuEx} - The {@link MenuEx} instance.
     * 2. {String} - The name of the selected menu item.
     * 3. {Integer} - The position of the selected menu item (e.g. 1 is the first menu item, 2 is the second, etc.).
     * 4. {Menu} - The `Menu` object.
     *
     * Additionally, if the menu was activated as a context menu event (see
     * {@link https://www.autohotkey.com/docs/v2/lib/GuiOnEvent.htm#Ctrl-ContextMenu}
     * and {@link https://www.autohotkey.com/docs/v2/lib/GuiOnEvent.htm#ContextMenu}), the following
     * parameters are also available:
     * 5. {Gui} - The `Gui` object
     * 6. {Gui.Control|String} - The `Gui.Control` object if one is associated with the context menu,
     *    else an empty string.
     * 7. {Integer|String} - The "Item" parameter as described in the documentation linked above.
     *
     * The function can also be the name of a class method. For details about this, see the section
     * "Extending MenuEx" in the description above {@link MenuEx}.
     *
     * The function's return value may be used if {@link MenuEx#ShowTooltips} is nonzero. For details
     * about how the return value is used, see {@link MenuEx.Prototype.SetTooltipHandler}.
     *
     * @param {String} [MenuItemOptions = ""] - Any options as described in
     * {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#Add}.
     *
     * @param {*} [HandlerTooltip] - See {@link MenuExItem.Prototype.SetTooltipHandler} for details about
     * this parameter.
     */
    __New(Name, CallbackOrSubmenu, MenuItemOptions := '', HandlerTooltip?) {
        this.__Name := Name
        this.__Value := CallbackOrSubmenu
        this.__Options := MenuItemOptions
        if IsSet(HandlerTooltip) {
            this.__HandlerTooltip := HandlerTooltip
        }
    }
    /**
     * Adds a checkbox next to the menu item.
     */
    Check() {
        this.MenuEx.Menu.Check(this.__Name)
    }
    /**
     * Deletes the menu item.
     */
    Delete() {
        this.MenuEx.Menu.Delete(this.__Name)
        this.MenuEx.__Item.Delete(this.__Name)
    }
    /**
     * Disables the menu item, causing the text to appear more grey than the surrounding text and
     * making it so the user cannot interact with the menu item.
     */
    Disable() {
        this.MenuEx.Menu.Disable(this.__Name)
    }
    /**
     * Enables the menu item, undoing the effect of {@link MenuExItem.Prototype.Disable} if it was
     * previously called.
     */
    Enable() {
        this.MenuEx.Menu.Enable(this.__Name)
    }
    /**
     * See {@link https://www.autohotkey.com/docs/v2/lib/Menu.htm#SetIcon} for details.
     */
    SetIcon(FileName, IconNumber := 1, IconWidth?) {
        this.MenuEx.Menu.SetIcon(this.__Name, FileName, IconNumber, IconWidth ?? unset)
    }
    /**
     * When {@link MenuEx#ShowTooltips} is true, there are three approaches for controlling the tooltip
     * text that is displayed when the user selects a menu item. When the user selects a menu item,
     * the return value returned by the function associated with the menu item is stored in a variable,
     * and the property {@link MenuExItem#HandlerTooltip} is evaluated to determine if a tooltip will be displayed,
     * and if so, what the text will be.
     *
     * Note that, by default, the value of {@link MenuExItem#HandlerTooltip} is an empty string.
     *
     * If {@link MenuExItem#HandlerTooltip} is an object, it is assumed to be a function or callable object.
     * The function is called with parameters:
     * 1. The {@link MenuEx} instance.
     * 2. The return value from the menu item's function.
     *
     * The function should return the string that will be displayed by the tooltip. If the function
     * returns an object or an empty string, no tooltip is displayed.
     *
     * If {@link MenuExItem#HandlerTooltip} is a significant string value, the return value from the menu
     * item's function is ignored and {@link MenuExItem#HandlerTooltip} is displayed in the tooltip.
     *
     * If {@link MenuExItem#HandlerTooltip} is zero or an empty string, and if the return value from the menu
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
        this.__HandlerTooltip := Value
    }
    /**
     * Toggles the display of a checkmark next to the menu item.
     */
    ToggleCheck() {
        this.MenuEx.Menu.ToggleCheck(this.__Name)
    }
    /**
     * Toggles the availability of the menu item.
     */
    ToggleEnable() {
        this.MenuEx.Menu.ToggleEnable(this.__Name)
    }
    /**
     * Removes a checkmark next to the menu item if one is present.
     */
    Uncheck() {
        this.MenuEx.Menu.Uncheck(this.__Name)
    }
    HandlerTooltip {
        Get => this.__HandlerTooltip
        Set {
            this.__HandlerTooltip := Value
        }
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

class MenuExItemCollection extends Map {
}

/**
 * Sets the global constant variables.
 *
 * @param {Boolean} [force = false] - When false, if `TreeViewEx_SetConstants` has already been called
 * (more specifically, if `tvex_flag_constants_set` has been set), the function returns immediately.
 * If true, the function executes in its entirety.
 */
MenuEx_SetConstants(force := false) {
    global
    if IsSet(MenuEx_flag_constants_set) && !force {
        return
    }
    ; https://learn.microsoft.com/en-us/windows/win32/menurc/keyboard-accelerator-messages
    WM_CHANGEUISTATE                := 0x0127
    WM_INITMENU                     := 0x0116
    WM_QUERYUISTATE                 := 0x0129
    WM_UPDATEUISTATE                := 0x0128

    UIS_CLEAR := 2
    UIS_INITIALIZE := 3
    UIS_SET := 1

    UISF_ACTIVE := 0x4
    UISF_HIDEACCEL := 0x2
    UISF_HIDEFOCUS := 0x1

    MenuEx_flag_constants_set := 1
}
