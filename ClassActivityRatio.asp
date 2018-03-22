<script runat=server language="vbscript">

    ' ############################################################################
    class ActivityRatio
        
        public ID
        public Name
        public BGColor
        public FGColor
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            Name = ""
            BGColor = ""
            FGColor = ""
        end sub
        
    end class
    
    
    ' ############################################################################
    ' READ ActivityRatio data
    public function getActivityRatio(id)
        dim SQL
        dim rst
        
        dim act
        set act = new ActivityRatio
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM ActivityRatio WHERE ID = " & id & " "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.Name = rst("Name")
            act.BGColor = rst("BGColor")
            act.FGColor = rst("FGColor")
            
        else
            Err.Raise 555, "ClassActivityRatio", "Ratio not found"
        end if
        rst.Close
        set rst = nothing
        
        set getActivityRatio = act
    end function
    
    
    public function getActivityRatios(Idioma)
        dim SQL
        dim rst
        dim arrRatios()
        dim stat, iRatio
        
        SQL = "SELECT id, name FROM ActivityRatio WHERE Idioma = '" & Idioma & "' ORDER BY id"
        set rst = Server.CreateObject("ADODB.RecordSet")
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        iRatio = 0
        while not rst.EOF
            redim preserve arrRatios(iRatio)
            
            set stat = getActivityRatio(rst("id"))
            
            set arrRatios(iRatio) = stat
            
            iRatio = iRatio + 1
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getActivityRatios = arrRatios
    end function
    
</script>