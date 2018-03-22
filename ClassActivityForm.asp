<%
    
    class ActivityFormResponse
        public idQuestion
        public idResponse
        
        public sub Class_Initialize()
            idQuestion = -1
            idResponse = -1
        end sub
    end class
    
    class ActivityForm
        
        public idActivity
        public idForm
        
        public numResponses
        public responses()
        
        public sub redim_responses(ind)
            redim preserve responses(ind)
        end sub
        
        public sub Class_Initialize()
            idActivity = -1
            idForm = -1
            numResponses = 0
        end sub
        
        ' Gets the response for a question
        public function getResponse(idQuestion)
            dim i, resResponse
            
            resResponse = -1
            if numResponses > 0 then
                for i = 0 to UBound(responses)
                    if responses(i).idQuestion = idQuestion then
                        resResponse = responses(i).idResponse
                    end if
                next
            end if
            
            getResponse = resResponse
        end function
        
    end class
    
    
    ' Loads the responses for the form in the activity
    public function loadActivityForm(idActivity)
        dim SQL, rst, rstQR
        dim actForm, Resp, nResp
        
        set actForm = new ActivityForm
        
        SQL = "SELECT idForm FROM Activity WHERE id = " & idActivity
        set rst = Server.CreateObject("ADODB.RecordSet")
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            actForm.idActivity = idActivity
            
            if NOT isNull(rst("idForm")) then
                
                actForm.idForm = rst("idForm")
                
                SQL = "SELECT idQuest, idResp " & _
                " FROM ActivityForm " & _
                " WHERE idActivity = " & idActivity & "  "
                set rstQR = Server.CreateObject("ADODB.RecordSet")
                rstQR.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
                nResp = 0
                while not rstQR.EOF
                
                    set Resp = new ActivityFormResponse
                    Resp.idQuestion = rstQR("idQuest")
                    Resp.idResponse = rstQR("idResp")
                    
                    ' Add the Response to the array of responses
                    actForm.redim_responses(nResp)
                    set actForm.responses(nResp) = Resp
                    
                    actForm.numResponses = actForm.numResponses + 1
                    
                    set Resp = nothing
                    nResp = nResp + 1
                    rstQR.MoveNext
                wend
                rstQR.Close
                set rstQR = nothing
                
            else
                ' No tiene formulario asignado!
                actForm.idForm = -1
            end if
        else
            Err.Raise 555, "ClassActivityForm", "Activity not found"
        end if
        
        rst.Close
        set rst = nothing
        
        set loadActivityForm = actForm
    end function
    
    
    
    public function saveActivityForm(Activity, arrResponses)
        dim r, SQL
        dim IDActivity, IDForm, IDBrand
        
        IDActivity = Activity.ID
        IDForm = Activity.idForm
        IDBrand = Activity.IDBrand
        
        
        ' *******************************************************************************************
        ' *******************************************************************************************
        ' Quizá no haya que borrar si no vienen respuestas, ya que puede tener datos históricos
        ' *******************************************************************************************
        ' *******************************************************************************************
        
        SQL = "DELETE FROM ActivityForm WHERE idActivity = " & idActivity
        ObjConnectionSQL.Execute SQL
        
        for each r in arrResponses
            
            SQL = "INSERT INTO ActivityForm (idActivity, idQuest, idResp, idForm, idBrand) " & _
            " VALUES ('" & idActivity & "', '" & r.IDQuestion & "', '" & r.IDResponse & "', '" & IDForm & "', '" & IDBrand & "')"
            ObjConnectionSQL.Execute SQL
            
        next
        
        
        recalculateKPIQuality Activity
        
    end function
    
    
    public function assignFormToActivity(idActivity, idForm)
        dim SQL
                
        SQL = "UPDATE Activity SET idForm = " & idForm & ", KPIQuality = NULL WHERE id = " & idActivity
        ObjConnectionSQL.Execute SQL
        
    end function
    
    
    public function formAppliesToActivity(Activity, Brand)
        dim yesno
        dim actWYear, actWMonth, actWHalf
        dim frm, frmDate
        dim DateBrand, DFormDate, DActDate
        
        yesno = false
        
        set frm = getForm(Brand.idForm)
        
        if frm.Enabled then
            frmDate = frm.DateFrom
            
            ' Regulariza la fecha con Q1 = día 1, Q2 = día 16
            if Day(frmDate) < 16 then
                ' Q1
                DFormDate = DateSerial(Year(FrmDate), Month(FrmDate), 1)
            else
                ' Q2
                DFormDate = DateSerial(Year(FrmDate), Month(FrmDate), 16)
            end if
            
            'Regulariza la fecha con Q1 = día 1, Q2 = día 16
            actWYear = Activity.WYear
            actWMonth = Activity.WMonth
            actWHalf = Activity.WHalf
            if actWHalf = 1 then
                DActDate = DateSerial(actWYear, actWMonth, 1)
            else
                DActDate = DateSerial(actWYear, actWMonth, 16)
            end if
            
            if DateDiff("d", DFormDate, DActDate) >= 0 then
                yesno = true
            else
                yesno = false
            end if
        else
            yesno = false
        end if
        
        
        formAppliesToActivity = yesno
    end function
    
    
    public function reassignActivityForm(idActivity, idNewForm)
        dim SQL
        
        ' Delete current responses (if any)
        SQL = "DELETE FROM ActivityForm WHERE idActivity = " & idActivity
        ObjConnectionSQL.Execute SQL
        
        ' Change IDForm from the Activity
        SQL = "UPDATE Activity SET idForm = " & idNewForm & ", KPIQuality = NULL WHERE ID = " & idActivity
        ObjConnectionSQL.Execute SQL
        
        
    end function
    
    
    public function activityFormCheckAndArrange(act, aBra, aCli)
        
        if aCli.activatedForms then
    
            dim frm, msgForm, showIdForm
            
            showIdForm = -1
            if CLng(act.ID) = -1 then
                'NEW Activity --> Show the QUALITY FORM if the brand is assigned to a form
                if aBra.idForm <> -1 then
                    if formAppliesToActivity(act, aBra) then
                        msgForm = "NEW ACTIVITY -- BRAND IS ASSIGNED TO FORM " & aBra.idForm
                        showIdForm = aBra.idForm
                    end if
                end if
            else
                'EDIT Activity
                
                dim actFrm
                
                msgForm = msgForm & "act.idForm " & act.idForm & " - aBra.idForm " & aBra.idForm & "<br>"
                
                if act.idForm <> aBra.idForm then
                    'The BRAND and the ACTIVITY have a DIFFERENT FORM
                    
                    msgForm = msgForm & "Brand i Activity tenen un form diferent<br>"
                    
                    if aBra.idForm <> -1 AND act.idForm = -1 then
                        ' The Brand has a form assigned, but NOT the activity
                        ' --> ASSIGN THE FORM TO THE ACTIVITY
                        
                        msgForm = msgForm & "L'Activity no té formulari, però el Brand si<br>"
                        
                        ' ASSIGN IF the date is after the FORM date
                        if formAppliesToActivity(act, aBra) then
                            set actFrm = loadActivityForm(act.ID)
                            assignFormToActivity act.ID, aBra.idForm
                            actFrm.idForm = aBra.idForm
                        
                            msgForm = msgForm & "Se li ha assignat perquè aplica idForm = [" & actFrm.idForm & "]<br>"
                            showIdForm = aBra.idForm
                        else
                            msgForm = msgForm & "Per les dates, no aplica el formulari<br>"
                        end if
                        
                    elseif aBra.idForm = -1 AND act.idForm <> -1 then
                        ' The Brand IS NOT assigned to a form, but the ACTIVITY IS ASSIGNED
                        ' Show the form with option to remove it
                        
                        msgForm = msgForm & "The brand doesn't have a form assigned any more. Unassign it?"
                        
                        showIdForm = act.idForm
                    
                    elseif aBra.idForm <> -1 AND act.idForm <> -1 then
                        ' The Brand and the Activity are assigned to a DIFFERENT form
                        ' REASSIGN IF the date is after the BRAND FORM date
                        if formAppliesToActivity(act, aBra) then
                            reassignActivityForm act.ID, aBra.idForm
                            showIdForm = aBra.idForm
                            msgForm = msgForm & "The form was reassigned"
                        else
                            showIdForm = act.idForm
                        end if
                        
                    else
                        set actFrm = loadActivityForm(act.ID)
                        if actFrm.idForm = -1 then
                            ' If the activity is not assigned to the form, assign it
                            if formAppliesToActivity(act, aBra) then
                                assignFormToActivity act.ID, aBra.idForm
                                actFrm.idForm = aBra.idForm
                                
                                msgForm = msgForm & "Form assigned to the activity"
                                showIdForm = aBra.idForm
                            end if
                        else
                            msgForm = msgForm & "FORM ASSIGNED " & actFrm.idForm
                            showIdForm = actFrm.idForm
                        end if
                    end if
                else
                    ' The BRAND and the ACTIVITY have the same FORM
                    
                    msgForm = msgForm & "Brand i Activity tenen EL MATEIX form<br>"
                    ' WILL SHOW THE FORM
                    if act.IDForm > -1 then
                        set actFrm = loadActivityForm(act.ID)
                        msgForm = msgForm & " idForm Assignat [" & act.idForm & "]<br>"
                        showIdForm = actFrm.IDForm
                    else
                        msgForm = msgForm & "No hi ha formulari assignat<br>"
                    end if
                end if
            end if
        end if
        
        ''''     Response.Write msgForm & "<br>"
        
        activityFormCheckAndArrange = showIDForm
    end function
  
  
    public function copyActivityForm(FromIdActivity, act)
        dim SQL
        dim aFromAct, aCli, aBra
        
        if act.idForm > -1 then
            ' If the activity had a form, delete it
            
            SQL = "UPDATE Activity SET idForm = NULL, KPIQuality = NULL WHERE ID = " & act.ID
            ObjConnectionSQL.Execute SQL

            SQL = "DELETE FROM ActivityForm WHERE idActivity = " & act.ID
            ObjConnectionSQL.Execute SQL
            
        end if
        
        set aFromAct = getActivity(FromIdActivity)
        if aFromAct.idForm > -1 then
            set aCli = getClient(act.IDClient)
            
            if aCli.activatedForms then
                set aBra = getBrand(act.IDBrand)
                if formAppliesToActivity(act, aBra) then
                    
                    SQL = "UPDATE Activity SET idForm = " & aFromAct.idForm & ", KPIQuality = NULL " & _
                    " WHERE ID = " & act.ID
                    ObjConnectionSQL.Execute SQL
                    
                    ' Inserts the same responses as the original form
                    SQL = "INSERT INTO ActivityForm (idActivity, idQuest, idResp, idForm, idBrand) " & _
                    " SELECT '" & act.ID & "', idQuest, idResp, idForm, idBrand " & _
                    " FROM ActivityForm " & _
                    " WHERE idActivity = " & FromIdActivity
                    ObjConnectionSQL.Execute SQL
                    
                end if
            end if
        end if
        
        
        ' Recalculate KPI quality
        recalculateKPIQuality act

    end function
    
    
    
    public function recalculateKPIQuality(Activity)
        dim SQL, rst
        dim kpiQuality
        
        kpiQuality = -1
        
        SQL = "SELECT SUM(fq.Weight * fr.RespValue) / 100 AS N " & _
        " FROM ActivityForm af " & _
        " INNER JOIN FormQuestion fq ON af.IDQuest = fq.IDQuest " & _
        " INNER JOIN FormResponse fr ON af.IDResp = fr.IDResp " & _
        " WHERE af.idActivity = " & Activity.ID
        set rst = Server.CreateObject("ADODB.RecordSet")
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            if not isNull(rst("N")) then
                kpiQuality = rst("N")
            end if
        end if
        set rst = nothing
        
        
        ' Actualiza la tabla Activity con el KPI calculado
        SQL = "UPDATE Activity SET KPIQuality = " & replace(kpiQuality,",",".") & " WHERE ID = " & Activity.ID
        ObjConnectionSQL.Execute SQL
        
        
    end function
   
   
%> 