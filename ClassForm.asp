<script runat="server" language="vbscript">

    ' ############################################################################
    class Form
        public ID
        public Name
        public indBaja
        public DateFrom
        public numQuestions
        public Questions  'Array of questions
        
        public sub Class_Initialize()
            ID = -1
            Name = ""
            indBaja = 0
            DateFrom = Date()
            numQuestions = 0
        end sub
        
        public property get Enabled
            Enabled = (indBaja=0)
        end property
        
        public property get weightTotal
            dim total, qst
            total = 0
            
            for each qst in Questions
                total = total + qst.Weight
            next
            
            weightTotal = total
        end property
        
        public property get canModify
            dim SQL, rst
            SQL = "SELECT COUNT(id) AS N FROM Activity WHERE idForm = " & idForm
            set rst = Server.CreateObject("ADODB.RecordSet")
            rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
            if rst("N") > 0 then
                canModify = false
            else
                canModify = true
            end if
            rst.Close
            set rst = nothing
            
        end property
        
        
    end class
    
    class FormQuestion
        public ID
        public IDForm
        public Text
        public Weight
        public Orden
        public IDRespType
        public Responses  'Array of responses
    
        public sub Class_Initialize()
            ID = -1
            Text = ""
            Orden = 9999
            IDRespType = 0
        end sub

        public property get numResponses
            dim total, rsp
            total = 0
            
            for each rsp in Responses
                total = total + 1
            next
            
            numResponses = total
        end property
        
    end class
    
    class FormResponse
        public ID
        public IDQuest
        public IDForm
        public Text
        public RespValue
        
        public sub Class_Initialize()
            ID = -1
            Text = ""
        end sub
        
    end class
    ' ############################################################################
    
    public function getForm(idForm)
        dim SQL, rst
        dim frm, rsp, qst
        dim iRsp, iQst
        
        set frm = new Form
        set rst = Server.CreateObject("ADODB.RecordSet")

        SQL = "SELECT * " & _
        " FROM Form " & _
        " WHERE idForm = " & idForm
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if NOT rst.EOF then
            
            set frm = new Form
            frm.ID = rst("IDForm")
            frm.Name = rst("Name")
            frm.indBaja = rst("indBaja")
            frm.DateFrom = rst("DateFrom")
            frm.Questions = getQuestions(idForm)
            
            on error resume next
            frm.numQuestions = UBound(frm.Questions) + 1
            on error goto 0
            
        else
            Err.Raise 555, "ClassForm", "Form not found"
        end if
        
        rst.Close
        set rst = nothing
        
        set getForm = frm
    end function
    
    
    ' *******************************************************************************
    ' Función para recoger las preguntas de un formulario y sus posibles respuestas
    public function getQuestions(idForm)
        dim SQL, rst, iQst, qst
        dim arrQst()
        
        set rst = Server.CreateObject("ADODB.RecordSet")

        SQL = "SELECT * " & _
        " FROM FormQuestion " & _
        " WHERE idForm = " & idForm & _
        " ORDER BY Orden "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        iQst = 0
        while not rst.EOF
            
            set qst = new FormQuestion
            
            qst.ID = rst("IDQuest")
            qst.IDForm = rst("IDForm")
            qst.Text = rst("Text")
            qst.Weight = rst("Weight")
            qst.Orden = rst("Orden")
            qst.IDRespType = rst("IDRespType")
            qst.Responses = getResponses(rst("IDQuest"))
            
            redim preserve arrQst(iQst)
            set arrQst(iQst) = qst
            
            iQst = iQst + 1
            rst.MoveNext
        wend
        rst.Close
        
        set rst = nothing
        
        getQuestions = arrQst
    end function
    

    ' *******************************************************************************
    ' Función para recoger las posibles respuestas a una pregunta de un formulario
    public function getResponses(idQuest)
        dim SQL, rst, iRsp, rsp
        dim arrRsp()
        
        set rst = Server.CreateObject("ADODB.RecordSet")

        SQL = "SELECT * " & _
        " FROM FormResponse " & _
        " WHERE idQuest = " & idQuest & _
        " ORDER BY idResp "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        iRsp = 0
        while not rst.EOF
            
            set rsp = new FormResponse
            
            rsp.ID = rst("IDResp")
            rsp.IDQuest = rst("IDQuest")
            rsp.IDForm = rst("IDForm")
            rsp.Text = rst("Text")
            rsp.RespValue = rst("RespValue")
            
            redim preserve arrRsp(iRsp)
            set arrRsp(iRsp) = rsp
            
            iRsp = iRsp + 1
            rst.MoveNext
        wend
        rst.Close
        
        set rst = nothing
        
        getResponses = arrRsp
    end function
    
    
    ' *******************************************************************************
    ' SAVE new or edit a FORM
    public function saveForm(frm)
        dim SQL, rst
        dim NewID
        dim isNew
        isNew = false
        
        if frm.ID > -1 then
            ' Not a new Form
            SQL = "UPDATE Form " & _
            " SET  " & _
            " Name = '" & replace(frm.Name, "'", "''") & "'" & _
            " , indBaja = '" & replace(frm.indBaja, "'", "") & "'" & _
            " , DateFrom = {d '" & Year(frm.DateFrom) & "-" & Right("0" & Month(frm.DateFrom), 2) & "-" & Right("0" & Day(frm.DateFrom), 2) & "'}" & _
            " WHERE idForm = " & frm.ID
        else
            isNew = true
            ' New Form
            SQL = "INSERT INTO Form (Name, indBaja, DateFrom) " & _
            " VALUES ('" & replace(frm.Name, "'", "''") & "', 0, {d '" & Year(frm.DateFrom) & "-" & Right("0" & Month(frm.DateFrom), 2) & "-" & Right("0" & Day(frm.DateFrom), 2) & "'} ) "
        end if
        
        ObjConnectionSQL.Execute SQL
        
        if frm.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassForm, saveForm", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            frm.ID = NewID
        end if
        
    end function
    
    public function deleteForm(idForm)
        dim SQL
        
        ' DELETE all forms related to activity
        SQL = "DELETE FROM ActivityForm WHERE IDForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
        ' DELETE all responses of the form
        SQL = "DELETE FROM FormResponse WHERE IDForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
        ' DELETE all questions of the form
        SQL = "DELETE FROM FormQuestion WHERE IDForm = " & idForm
        ObjConnectionSQL.Execute SQL

        ' DELETE the FORM
        SQL = "DELETE FROM Form WHERE IDForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
        ' UNASSIGN the form from the brands
        SQL = "UPDATE Brand SET idForm = NULL WHERE idForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
    end function
    
    public function disableForm(idForm)
        dim SQL
        
        SQL = "UPDATE Form SET indBaja = 1 WHERE idForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
    end function

    public function enableForm(idForm)
        dim SQL
        
        SQL = "UPDATE Form SET indBaja = 0 WHERE idForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
    end function


    ' *******************************************************************************
    ' SAVE new or edit a QUESTION
    public function saveQuestion(qst)
        dim SQL, rst
        dim NewID
        dim isNew
        isNew = false
        
        if qst.ID > -1 then
            ' Not a new Question
            SQL = "UPDATE FormQuestion " & _
            " SET  " & _
            " Text = '" & replace(qst.Text, "'", "''") & "'" & _
            " , Weight = '" & replace(replace(qst.Weight,",","."),"'","") & "'" & _
            " , Orden = '" & replace(qst.Orden,"'","") & "'" & _
            " , IDRespType = '" & replace(qst.IDRespType,"'","") & "' " & _
            " WHERE idQuest = " & qst.ID
        else
            isNew = true
            ' New Question
            SQL = "INSERT INTO FormQuestion (IDForm, Text, Weight, Orden, IDRespType) " & _
            " VALUES ( " & qst.IDForm & " " & _
            " , '" & replace(qst.Text, "'", "''") & "' " & _
            " , '" & replace(replace(qst.Weight,",","."),"'","") & "' " & _
            " , '" & replace(qst.Orden,"'","") & "' " & _
            " , '" & replace(qst.IDRespType,"'","") & "' " & _
            " ) "
        end if
        
        ObjConnectionSQL.Execute SQL
        
        if qst.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassForm, saveQuestion", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            
            ' Al crear una pregunta nueva, debería reordenar
            
            qst.ID = NewID
        end if
        
    end function
    
    public function deleteQuestion(idQuest)
        dim SQL
        
        ' DELETE all responses of the question
        SQL = "DELETE FROM FormResponse WHERE IDQuest = " & idQuest
        ObjConnectionSQL.Execute SQL
        
        ' DELETE all questions of the question
        SQL = "DELETE FROM FormQuestion WHERE IDQuest = " & idQuest
        ObjConnectionSQL.Execute SQL
        
    end function
    

    ' *******************************************************************************
    ' SAVE new or edit a RESPONSE
    public function saveResponse(rsp)
        dim SQL, rst
        dim NewID
        dim isNew
        isNew = false
        
        if rsp.ID > -1 then
            ' Not a new Response
            SQL = "UPDATE FormResponse " & _
            " SET  " & _
            " Text = '" & replace(rsp.Text, "'", "''") & "'" & _
            " , RespValue = '" & replace(replace(rsp.RespValue,",","."),"'","") & "'" & _
            " WHERE idResp = " & rsp.ID
        else
            isNew = true
            ' New Response
            SQL = "INSERT INTO FormResponse (IDQuest, IDForm, Text, RespValue) " & _
            " VALUES ( " & rsp.IDQuest & " " & _
            " , " & rsp.IDForm & " " & _
            " , '" & replace(rsp.Text, "'", "''") & "' " & _
            " , '" & replace(replace(rsp.RespValue,",","."),"'","") & "' " & _
            " ) "
        end if
        
        ObjConnectionSQL.Execute SQL
        
        if rsp.ID < 0 then
            ' Si era nueva, busca el nuevo ID y lo informa en la clase
            SQL = "SELECT @@IDENTITY"
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NewID = rst.Fields(0)
            else
                Err.Raise 1, "ClassForm, saveResponse", "Cannot read @@IDENTITY"
            end if
            rst.Close
            set rst = nothing
            
            rsp.ID = NewID
        end if
        
    end function

    public function deleteResponse(idResp)
        dim SQL
        
        ' DELETE all responses of the question
        SQL = "DELETE FROM FormResponse WHERE IDResp = " & idResp
        ObjConnectionSQL.Execute SQL
        
    end function
    
    
    public sub removeFormFromBrand(idBrand, idForm, deleteHistory)
        ' deleteHistory is boolean
        dim SQL
        
        SQL = "UPDATE Brand SET idForm = NULL WHERE idBrand = " & idBrand
        ObjConnectionSQL.Execute SQL
        
        if deleteHistory then
            SQL = "DELETE FROM ActivityForm WHERE idBrand = " & idBrand
            ObjConnectionSQL.Execute SQL

            SQL = "UPDATE Activity SET idForm = NULL WHERE idBrand = " & idBrand & " AND idForm = " & idForm
            ObjConnectionSQL.Execute SQL
        end if
        
    end sub
    
    
    public sub removeFromPromotions(idBrand, idForm)
        dim SQL
        
        SQL = "DELETE FROM ActivityForm WHERE idBrand = " & idBrand & " AND idForm = " & idForm
        ObjConnectionSQL.Execute SQL
        
        SQL = "UPDATE Activity SET idForm = NULL WHERE idForm = " & idForm & " AND idBrand = " & idBrand
        ObjConnectionSQL.Execute SQL
        
    end sub
    
    
    public sub assignBrand(idBrand, idForm)
        dim SQL
        
        SQL = "UPDATE Brand SET idForm = " & idForm & " WHERE idBrand = " & idBrand
        ObjConnectionSQL.Execute SQL
        
    end sub
    
    
    public sub reassignBrand(idBrand, idForm, deleteHistory)
        dim SQL
        
        SQL = "UPDATE Brand SET idForm = " & idForm & " WHERE idBrand = " & idBrand
        ObjConnectionSQL.Execute SQL
        
        if deleteHistory then
            SQL = "DELETE FROM ActivityForm WHERE idBrand = " & idBrand
            ObjConnectionSQL.Execute SQL
            
            SQL = "UPDATE Activity SET idForm = NULL WHERE idBrand = " & idBrand & " AND idForm = " & idForm
            ObjConnectionSQL.Execute SQL
        end if
        
    end sub
    

    public function recalculateFormKPIQuality(idForm)
        dim SQL, rst
        dim act

        set rst = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT ID FROM Activity " & _
        " WHERE idForm = " & idForm
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        while not rst.EOF
            
            set act = getActivity(rst("ID"))
            recalculateKPIQuality act
            set act = nothing
            
            rst.MoveNext
        wend
        rst.Close
        set rst = nothing
        
    end function
    
    
</script>