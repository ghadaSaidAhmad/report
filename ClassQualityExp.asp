<script runat=server language="vbscript">

class QExp
    public ID
    public Descripcion
    public Valoracion
    public Selected
end class

Class QualityExp
    
    dim arrQuality()
    
    
    public sub Class_Initialize()
        dim rst, SQL, strOut
            
        set rst = Server.CreateObject("ADODB.RecordSet")

        SQL = "SELECT ID, Descripcion, Valoracion " & _
        " FROM CalidadExp " & _
        " ORDER BY ID "
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        dim nQ: nQ = 0
        dim newQ
        while not rst.EOF
            
            set newQ = new QExp
            newQ.ID = rst("ID")
            newQ.Descripcion = rst("Descripcion")
            newQ.Valoracion = rst("Valoracion")
            
            redim preserve arrQuality(nQ)
            set arrQuality(nQ) = newQ
            
            nQ = nQ + 1
            rst.MoveNext
        wend
        rst.Close

    end sub
    
    
    Function getQualitySelectors()
        
        getQualitySelectors = ""
    End Function



End Class

</script>