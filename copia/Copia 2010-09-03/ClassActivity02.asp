<script runat=server language="vbscript">

    ' ############################################################################
    class Activity02
        
        public ID
        public IDBrand
        public IDClient
        public WYear
        public WMonth
        public WHalf
        public IDType
        public Name
        public LastUpdatedBy
        public LastUpdatedDate
        public NShops
        public PercentComplaint
        public Status
        public IDStatus
        public StatusBGColor
        public StatusFGColor
        
        'Init method
        public sub Class_Initialize()
            ID = -1
            IDBrand = 0
            IDClient = 0
            WYear = 1900
            WMonth = 1
            WHalf = 1
            IDType = 0
            Name = ""
            'NShops = 0
            'PercentComplaint = 0.0
            Status = ""
            IDStatus = 1
        end sub


        ' Texto que se muestra en el GRID
        public property get GridText
            
            GridText = Name
            
        end property
        

        public property get TextBGColor
            dim ret
            ret = ""
            
            if GridText<>"" then
                ret = StatusBGColor
            else
                ret = ""
            end if

            TextBGColor = ret
        end property
        

        public property get TextFGColor
            dim ret
            ret = ""
            
            if GridText<>"" then
                ret = StatusFGColor
            else
                ret = ""
            end if

            TextFGColor = ret
        end property

    end class
    
    
    ' ############################################################################
    ' READ Activity data
    public function getActivity02(id)
        dim SQL
        dim rst
        
        dim act
        set act = new Activity02
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT act.*, CONVERT(varchar, act.LastUpdatedDate, 103) + ' ' + CONVERT(varchar, act.LastUpdatedDate, 108) AS UpdatedDate, " & _
        " act2.NShops, act2.PercentComplaint, " & _
        " act2.IDStatus, st.Name AS StatusName, st.FGColor AS StatusFGColor, st.BGColor AS StatusBGColor, " & _
        " em.ApellidosNombre AS NombreUpdated " & _
        " FROM Activity act " & _
        " INNER JOIN Activity02 act2 ON act.id = act2.id " & _
        " LEFT JOIN EmpleadosGlobal em ON act.LastUpdatedBy = em.IDEmpleado " & _
        " LEFT JOIN ActivityStatus st ON act2.IDStatus = st.ID " & _
        " WHERE act.ID = " & id
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            act.ID = id
            act.IDBrand = rst("IDBrand")
            act.IDClient = rst("IDClient")
            act.WYear = rst("WYear")
            act.WMonth = rst("WMonth")
            act.WHalf = rst("WHalf")
            act.IDType = rst("IDType")
            act.Name = rst("Name")
            act.LastUpdatedBy = rst("NombreUpdated")
            act.LastUpdatedDate = rst("UpdatedDate")
            
            act.NShops = rst("NShops")
            act.PercentComplaint = rst("PercentComplaint")

            act.IDStatus = rst("IDStatus")
            act.Status = rst("StatusName")
            act.StatusBGColor = rst("StatusBGColor")
            act.StatusFGColor = rst("StatusFGColor")
            
        else
            Err.Raise 555, "ClassActivity02", "Activity not found " & SQL
        end if
        rst.Close
        set rst = nothing
        
        set getActivity02 = act
    end function
    
    
    
    ' ############################################################################
    public function getActivities02(WYear, WMonth, IDClient, IDBrand, IDType)
        dim SQL
        dim rst
        dim act1, act2
        dim arrActivity(1) ' Array con una actividad por quincena
        
        
        set act1 = new Activity02
        set act2 = new Activity02
        
        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID, WHalf FROM Activity " & _
        " WHERE WYear = " & WYear & " AND WMonth = " & WMonth & _
        " AND IDClient = " & IDClient & " AND IDBrand = " & IDBrand & _
        " AND IDType = " & IDType
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            if rst("WHalf") = 1 then
                set act1 = getActivity02(rst("id"))
            elseif rst("WHalf") = 2 then
                set act2 = getActivity02(rst("id"))
            else
                'Pero cuantas quincenas tiene un mes??
            end if
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
        set arrActivity(0) = act1
        set arrActivity(1) = act2
        
        getActivities02 = arrActivity
    end function
    
    

    ' ############################################################################
    'SAVE new or edit
    public sub saveActivity02(act)
        dim SQL
        dim rst
        dim NewID
        dim isNew
        isNew = false
        
        if act.ID > -1 then
            ' Not a new Activity
            SQL = "UPDATE Activity " & _
            " SET  " & _
            " Name = '" & replace(act.Name, "'", "''") & "'" & _
            " , LastUpdatedBy = '" & replace(session("IDEmpleado"), "'", "''") & "'" & _
            " , LastUpdatedDate = GETDATE() " & _
            " WHERE id = " & act.ID
            
        else
            isNew = true
            ' New Activity
            SQL = "INSERT INTO Activity (IDBrand, IDClient, WYear " & _
            " , WMonth, WHalf, IDType, Name " & _
            " , LastUpdatedBy, LastUpdatedDate " & _
            " ) " & _
            " VALUES (" & act.IDBrand & ", " & act.IDClient & ", " & act.WYear & " " & _
            " , " & act.WMonth & ", " & act.WHalf & ", " & act.IDType & ", '" & Replace(act.Name, "'", "''") & "' " & _
            " , " & session("IDEmpleado") & ", GETDATE() " & _
            " )"
        
        end if
        
        ObjConnectionSQL.Execute SQL
        
        if act.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassActivity02", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            act.ID = NewID
        end if
        
        
        act.NShops = Replace(act.NShops, ",", ".")
        act.PercentComplaint = Replace(act.PercentComplaint, ",", ".")
        if act.NShops<>"" then if NOT isNumeric(act.NShops) then act.NShops = 0
        if act.PercentComplaint<>"" then if NOT isNumeric(act.PercentComplaint) then act.PercentComplaint = 0
        
        dim sNShops: sNShops = "NULL"
        if act.NShops <> "" then sNShops = act.NShops
        dim sPercentComplaint: sPercentComplaint = "NULL"
        if act.PercentComplaint <> "" then sPercentComplaint = act.PercentComplaint
        
        if NOT isNew then
            
            SQL = "UPDATE Activity02 " & _
            " SET " & _
            " NShops = " & sNShops & " " & _
            " , PercentComplaint = " & sPercentComplaint & " " & _
            " , IDStatus = '" & replace(act.IDStatus, "'", "") & "' " & _
            " WHERE id = " & act.ID
            
            
        else
            
            SQL = "INSERT INTO Activity02 (" & _
            " ID " & _
            " , NShops " & _
            " , PercentComplaint " & _
            " , IDStatus " & _
            " ) " & _
            " VALUES (" & _
            " " & act.ID & " " & _
            " , " & sNShops & " " & _
            " , " & sPercentComplaint & " " & _
            " , '" & Replace(act.IDStatus, "'", "") & "' " & _
            " )"
            
        end if
        

        ObjConnectionSQL.Execute SQL
        
    end sub
    


    ' ############################################################################
    sub deleteActivity02(id)
        
        dim SQL
        SQL = "DELETE FROM Activity WHERE ID = " & ID
        ObjConnectionSQL.Execute SQL
        
        SQL = "DELETE FROM Activity02 WHERE ID = " & ID
        ObjConnectionSQL.Execute SQL

    end sub

</script>