function Pagina(id){

                                    var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                   
           
                                    var result;
                                    var i = 0;
                                    var ar = new Array();
                                    while (!RS.EOF) {
                                        if(id == RS.fields(0).value){
                                             document.getElementById("id_edit_user").value = RS.fields(0).value;
                                             document.getElementById("firstName1").value = RS.fields(1).value;
                                             document.getElementById("lastName1").value = RS.fields(2).value;
                                             document.getElementById("btn_novo").style.display="none";
                                             document.getElementById("btn_editar").style.display="block";
                                        }      
                                              RS.MoveNext();
                                    }
  }



 

  function QuantidadeRegistros(){

                                    var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);

                                    var i = 0;
                                    while (!RS.EOF) {
                                      i++;
                                       RS.MoveNext();
                                    }
                                        
      

      return i;
 
  }

  function Listar(){
			var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
			var i = 0;
			var ar = new Array();
			while (!RS.EOF) {
	     	i++;
				ar[i] =  "<li class='list-group-item d-flex justify-content-between'>"+
                 "<input type='hidden' id='"+RS.fields(0).value +"' name='id_usu'"+RS.fields(0).value +"' value='"+RS.fields(0).value +"'>" 
                  +RS.fields(1).value+" "+ RS.fields(2).value+"<button type='button' class='btn btn-link' onclick='Delete("+RS.fields(0).value +");'>Excluir</button>"
						    +"<button type='button' class='btn btn-link' data-toggle='modal' data-target='#myModal' id='"+RS.fields(0).value +"' onclick='Pagina("+RS.fields(0).value +")'>Editar </button></li>";
                
	            RS.MoveNext();
      }
        	
			var result  = ar;

    return result;
  }

function ListarBustaAutocomplete(){

      var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);

      var i = 0;
      var ar = new Array();
      while (!RS.EOF) {
              i++;
              ar[i] =  RS.fields(1).value+ RS.fields(2).value;
              RS.MoveNext();
      }
          
      var result  = ar;

    return result;
}




  function Insert(nome,sobrenome){
			var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                  
                                    var sql = "INSERT INTO usuario(UserName,LastName)VALUES('"+nome+"','"+sobrenome+"')";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                    alert("Registro inserido com sucesso!");
                                    RS.close();
  }



  function Update(id,nome,subrenome){


                                    var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "UPDATE usuario SET UserName = '"+nome+"',LastName = '"+subrenome+"' WHERE id = "+id;
                                  
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                    alert("Registro alterado com sucesso");
                                    RS.close();



  }
function Delete(id){
			                              var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    	
                                    var sql = "DELETE FROM usuario  WHERE id = "+ id;
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                    alert("Registro excluido com sucesso!");
                                    RS.close();
  }

function protegercodigoMouse() {
    if (event.button==2||event.button==3){
        alert('Codigo protegido!');}
    }