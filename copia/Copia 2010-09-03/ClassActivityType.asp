<script runat=server language="vbscript">

    ' ############################################################################
    class ActivityType
        
        public ID
        public Name
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            Name = "Not Set"
        end sub
        
    end class
    
    
    ' ############################################################################
    ' READ ActivityType data
    public function getActivityType(id, idioma)
        dim SQL
        dim rst
        
        dim act
        set act = new ActivityType
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM ActivityType WHERE ID = " & id & " AND Idioma = '" & idioma & "' "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.Name = rst("Name")
            
        else
            Err.Raise 555, "ClassActivityType", "Type not found"
        end if
        rst.Close
        set rst = nothing
        
        set getActivityType = act
    end function
    
    
    public function getActivityTypes(idioma)
        dim SQL
        dim rst
        dim arrTypes()
        dim typ, iType
        
        SQL = "SELECT id, name FROM ActivityType WHERE Idioma = '" & idioma & "' ORDER BY id"
        set rst = Server.CreateObject("ADODB.RecordSet")
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        iType = 0
        while not rst.EOF
            redim preserve arrTypes(iType)
            
            set typ = getActivityType(rst("id"), idioma)
            
            set arrTypes(iType) = typ
            
            iType = iType + 1
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getActivityTypes = arrTypes
    end function
    
</script>