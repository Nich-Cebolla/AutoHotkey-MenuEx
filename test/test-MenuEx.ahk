
#include ..\src\MenuEx.ahk

test()

class test {
    static Call() {
        g := this.g := Gui('+Resize')
        tv := this.tv := g.Add('TreeView', 'w600 r20 vTv')
        m := this.m := MapEx()
        i := 0
        loop 2 {
            ++i
            id := tv.add('test' i)
            m.set(id, MapEx())
            m2 := m.Get(id)
            loop 10 {
                _id := tv.Add('test' i '-' A_Index, id)
                m2.Set(_id, MapEx())
            }
        }
        o := this.o := {
            CaseSense: false
          , TooltipHandler: ''
          , SelectionHandler: ''
          , ShowTooltips: true
          , WhichMethod: 1
          , TooltipDefaultOptions: ''
          , ItemVisibilityHandler: ''
          , IsMenuBar: false
        }
        mex := this.mex := MenuEx(o)
        mex.Add('callback1', callback1)
        mex.Add('callback2', callback2)
        tv.OnEvent('ContextMenu', mex)
        g.Show()
    }
}


class MapEx extends Map {
    __New() {
        this.CaseSense := false
    }
}

callback1(MenuExObj, Params, Output := true) {
    s := ''
    for prop, val in params.OwnProps() {
        if prop == 'Token' {
            ind := '    '
            s .= prop ' :: {`n'
            for _prop, _val in val.OwnProps() {
                s .= ind _prop ' :: ' (IsObject(_val) ? '{ ' Type(_val) ' }' : _val) '`n'
            }
            s .= '}`n'
        } else {
            s .= prop ' :: ' (IsObject(val) ? '{ ' Type(val) ' }' : val) '`n'
        }
    }
    if Output {
        OutputDebug(s)
    }
    return s
}

callback2(MenuExObj, Params) {
    s := callback1(MenuExObj, Params, false)
    OutputDebug(s)
    return s
}
