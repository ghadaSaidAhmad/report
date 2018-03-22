<script runat=server language="vbscript">

    ' ############################################################################
    class ActivityStatus
        
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
    ' READ ActivityStatus data
    public function getActivityStatus(id)
        dim SQL
        dim rst
        
        dim act
        set act = new ActivityStatus
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT * FROM ActivityStatus WHERE ID = " & id & " "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.Name = rst("Name")
            act.BGColor = rst("BGColor")
            act.FGColor = rst("FGColor")
            
        else
            Err.Raise 555, "ClassActivityStatus", "Status not found"
        end if
        rst.Close
        set rst = nothing
        
        set getActivityStatus = act
    end function
    
    
    public function getActivityStatuses()
        dim SQL
        dim rst
        dim arrStatuses()
        dim stat, iStatus
        
        SQL = "SELECT id, name FROM ActivityStatus ORDER BY id"
        set rst = Server.CreateObject("ADODB.RecordSet")
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        iStatus = 0
        while not rst.EOF
            redim preserve arrStatuses(iStatus)
            
            set stat = getActivityStatus(rst("id"))
            
            set arrStatuses(iStatus) = stat
            
            iStatus = iStatus + 1
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        getActivityStatuses = arrStatuses
    end function
    
</script>