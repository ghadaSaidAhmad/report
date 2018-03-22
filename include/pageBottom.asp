<script language="javascript">
    <%if bottomMessage<>"" then %>
        alert('<%=replace(replace(bottomMessage, "\", "\\"), "'", "\'") %>');
    <%end if %>
    
    <%if refreshOpener then %>
        try{
            window.opener.thisForm.submit();
        }catch(e){}
    <%end if %>
    
    <%if showMenu then %>
        try{
            showMenu();
        }catch(e){}
    <%end if %>
    
    try{
        DIV_WAIT.style.display = 'none';
    }catch(e){}
    
    window.focus();
</script>


