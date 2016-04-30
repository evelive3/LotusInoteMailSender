<font face="Default Sans Serif,Verdana,Arial,Helvetica,sans-serif" size="2">
    <table cellspacing="0" cellpadding="2"
           style="border:1px solid #009a63;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:small;"
           width="100%" align="center">
        <tbody>
        <tr>
            {% for row in rows.keys() %}
            <td width="auto" style="text-align: center; border: 1px solid #009a63; background: #009a63"><p style="color: #ffffff">{{ row }}</p></td>
            {% endfor %}
        </tr>
        <tr>
            {% for row in rows.values() %}
            <td width="auto" style="text-align: center; border: 1px solid #009a63;"><p style="color: #009a63">{{ row }}</p></td>
            {% endfor %}
        </tr>
        </tbody>
    </table>
    <div></div>
</font>