

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<div class="wrapper">
    <div class="content">
        
        <%
        sids = request("sids")
        splitsids = split(sids, ",")

        for i = 0 TO UBOUND(splitsids)
            response.Write "<br> Nyhed " & splitsids(i)

            oConn.execute("INSERT INTO systemmeddelelser_rel SET sysmed = "& splitsids(i) & ", medid = "& session("mid"))

        next
        %>

    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->