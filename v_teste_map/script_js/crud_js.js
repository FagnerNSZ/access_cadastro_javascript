  function Pagina(id){
    //  console.log("location.href=C:\\v_teste_map\\index.html?cod = "+id);
    //location.href=C:\\v_teste_map\\index.html?cod = 1
    //var quebra = variaveis[1].split("=");
    //console.log(quebra);
    //window.location = "?id="+id;
    //document.getElementById("firstName1Edit").value,document.getElementById("lastName1Edit").value);
   
   // Update(id,document.getElementById("firstName1Edit").value,document.getElementById("lastName1Edit").value);

  //window.location = "?id="+id;
  var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                   
           
      var result;
      var i = 0;
      var ar = new Array();
      //window.location = "?id="+id;
      while (!RS.EOF) {
            //i++;
          if(id == RS.fields(0).value){

           // window.open("C:\\v_teste_map\\index.html?id="+id);
               
               
               console.log(RS.fields(0).value);  
               document.getElementById("id_edit_user").value = RS.fields(0).value;
               document.getElementById("firstName1").value = RS.fields(1).value;
               document.getElementById("lastName1").value = RS.fields(2).value;
               document.getElementById("btn_novo").style.display="none";
               document.getElementById("btn_editar").style.display="block";


             // return RS.fields(0).value;

//window.location = "?id="+id;
                /*result = '<div class="modal" id="myModal">'+
                '<div class="modal-dialog">'+
      '<div class="modal-content"> '+
        '<div class="modal-header">'+
          '<h4 class="modal-title">Editar usuário</h4>'+
          '<button type="button" class="close" data-dismiss="modal">&times;</button>'+
        '</div>'+
        '<div class="modal-body">' +
          '<form class="needs-validation" novalidate>'+
            '<div class="row">'+
              '<div class="col-md-6 mb-3">'+
                '<label for="firstName">Nome</label>'+
                '<input type="hidden" id="'+id+'" name="custId" value="'+id+'">'+
                '<input type="text" class="form-control" id="firstName1Edit" placeholder="" value="'+RS.fields(0).value+'" required>'+
                '<input type="text" class="form-control" id="firstName1Edit" placeholder="" value="'+RS.fields(1).value+'" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio.'+
                '</div>'+
              '</div>'+
              '<div class="col-md-6 mb-3">'+
                '<label for="lastName">Sobrenome</label>'+
                '<input type="text" class="form-control" id="lastName1Edit" placeholder="" value="'+RS.fields(2).value+'" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio'+
                '</div>'+
              '</div>'+
            '</div>'+
            '<hr class="mb-4">'+
            '<button class="btn btn-primary btn-lg btn-block"  onclick="teste();">Update</button>'+
            '<button class="btn btn-primary btn-lg btn-block"  onclick="Update(27,document.getElementById("firstName1Edit").value,document.getElementById("lastName1Edit").value);">Salvar</button>'+
            '<button type="button" class="btn btn-link" data-toggle="modal" data-target="#myModal" onclick="Pagina('+RS.fields(0).value+')";>Bookmark </button>'+
        '</form>'+
        '</div>'+
        '<div class="modal-footer">'+
          '<button type="button" class="btn btn-danger" data-dismiss="modal">Close</button>'+
        '</div>'+
      '</div>'+
    '</div>';*/
       //result =RS.fields(0).value;
          }      

                RS.MoveNext();

      }
       
      //return result;
  }



 function ListInModal(){
  
      var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);




           // alert(RS.Fields.Item("UserName")+ " "+RS.Fields.Item("LastName") );
         
           /*alert(RS.Fields.Item("ID")+ 
             RS.Fields.Item("UserName") +
             RS.Fields.Item("LastName")+ 
             RS.Fields.Item("name"));
             */

         /*while (!RS.EOF) {
            console.log(RS.Fields.Item("UserName"));
            //rs.fields(0).value 
                var result = "<li class='list-group-item d-flex justify-content-between'>"+
                      "<span>"+RS.Fields.Item("UserName")+"</span>"+"<strong>"+RS.Fields.Item("LastName")+"</strong></li>";
        }*/

      var i = 0;
      var ar = new Array();
      while (!RS.EOF) {
        //console.log( + RS.fields(0).value + RS.fields(1).value+ RS.fields(2).value);
        //if(RS.fields(0).value == 27){
        i++;
        /*ar[i] =  "<li class='list-group-item d-flex justify-content-between'>"+
                 "<input type='hidden' id='"+RS.fields(0).value +"' name='id_usu"+RS.fields(0).value +"' value='"+RS.fields(0).value +"'>" 
                 +RS.fields(0).value+ RS.fields(1).value+ RS.fields(2).value+"<button type='button' class='btn btn-link' onclick='Delete("+RS.fields(0).value +");'>Excluir</button>"
                +"<button type='button' class='btn btn-link' data-toggle='modal' data-target='#myModal' id='"+RS.fields(0).value +"'>Editar </button></li>'";
                */

//<div class="modal" id="myModal">
    
  //</div>
                ar[i] = '<div class="modal" id="myModal">'+
                '<div class="modal-dialog">'+
      '<div class="modal-content"> '+
        '<div class="modal-header">'+
          '<h4 class="modal-title">Editar usuário</h4>'+
          '<button type="button" class="close" data-dismiss="modal">&times;</button>'+
        '</div>'+
        '<div class="modal-body">' +
          '<form class="needs-validation" novalidate>'+
            '<div class="row">'+
              '<div class="col-md-6 mb-3">'+
                '<label for="firstName">Nome</label>'+
                '<input type="hidden" id="custId" name="custId" value="'+RS.fields(0).value+'">'+
                '<input type="text" class="form-control" id="firstName1Edit" placeholder="" value="'+RS.fields(0).value+'" required>'+
                '<input type="text" class="form-control" id="firstName1Edit" placeholder="" value="'+RS.fields(1).value+'" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio.'+
                '</div>'+
              '</div>'+
              '<div class="col-md-6 mb-3">'+
                '<label for="lastName">Sobrenome</label>'+
                '<input type="text" class="form-control" id="lastName1Edit" placeholder="" value="'+RS.fields(2).value+'" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio'+
                '</div>'+
              '</div>'+
            '</div>'+
            '<hr class="mb-4">'+
            '<button class="btn btn-primary btn-lg btn-block"  onclick="teste();">Update</button>'+
            '<button class="btn btn-primary btn-lg btn-block"  onclick="Update(27,document.getElementById("firstName1Edit").value,document.getElementById("lastName1Edit").value);">Salvar</button>'+
            '<button type="button" class="btn btn-link" data-toggle="modal" data-target="#myModal" onclick="Pagina('+RS.fields(0).value+')";>Bookmark </button>'+
        '</form>'+
        '</div>'+
        '<div class="modal-footer">'+
          '<button type="button" class="btn btn-danger" data-dismiss="modal">Close</button>'+
        '</div>'+
      '</div>'+
    '</div>';
       // }
              RS.MoveNext();

      }
          
      var result  = ar[i];


      /* '<div class="modal-dialog">'+
      '<div class="modal-content"> '+
        '<div class="modal-header">'+
          '<h4 class="modal-title">Editar usuário</h4>'+
          '<button type="button" class="close" data-dismiss="modal">&times;</button>'+
        '</div>'+
        '<div class="modal-body">' +
          '<form class="needs-validation" novalidate>'+
            '<div class="row">'+
              '<div class="col-md-6 mb-3">'+
               '<input type="hidden" id="id_edit_user" name="custId" value="'+RS.fields(0).value+'">'+
                '<label for="firstName">Nome</label>'+
                '<input type="text" class="form-control" id="firstName1Edit" placeholder="" value="'+RS.fields(0).value+sql+'" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio.'+
                '</div>'+
              '</div>'+
              '<div class="col-md-6 mb-3">'+
                '<label for="lastName">Sobrenome</label>'+
                '<input type="text" class="form-control" id="lastName1Edit" placeholder="" value="" required>'+
                '<div class="invalid-feedback">'+
                  'Este campo é obrigatorio'+
                '</div>'+
              '</div>'+
            '</div>'+
            '<hr class="mb-4">'+
            '<button class="btn btn-primary btn-lg btn-block"  onclick="Insert(document.getElementById("firstName1").value,document.getElementById("lastName1").value);">Salvar</button>'+
        '</form>'+
        '</div>'+
        '<div class="modal-footer">'+
          '<button type="button" class="btn btn-danger" data-dismiss="modal">Close</button>'+
        '</div>'+
      '</div>'+
    '</div>';
*/
    return result;
  }


  function Listar(){
			var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    var sql = "SELECT  * FROM  usuario";
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);




           // alert(RS.Fields.Item("UserName")+ " "+RS.Fields.Item("LastName") );
         
           /*alert(RS.Fields.Item("ID")+ 
             RS.Fields.Item("UserName") +
             RS.Fields.Item("LastName")+ 
             RS.Fields.Item("name"));
             */

				 /*while (!RS.EOF) {
				    console.log(RS.Fields.Item("UserName"));
				    //rs.fields(0).value 
		            var result = "<li class='list-group-item d-flex justify-content-between'>"+
	 			              "<span>"+RS.Fields.Item("UserName")+"</span>"+"<strong>"+RS.Fields.Item("LastName")+"</strong></li>";
				}*/

			var i = 0;
			var ar = new Array();
			while (!RS.EOF) {
			 	//console.log( + RS.fields(0).value + RS.fields(1).value+ RS.fields(2).value);
	     	i++;
				ar[i] =  "<li class='list-group-item d-flex justify-content-between'>"+
                 "<input type='hidden' id='"+RS.fields(0).value +"' name='id_usu'"+RS.fields(0).value +"' value='"+RS.fields(0).value +"'>" 
                 +RS.fields(0).value+ RS.fields(1).value+ RS.fields(2).value+"<button type='button' class='btn btn-link' onclick='Delete("+RS.fields(0).value +");'>Excluir</button>"
						    +"<button type='button' class='btn btn-link' data-toggle='modal' data-target='#myModal' id='"+RS.fields(0).value +"' onclick='Pagina("+RS.fields(0).value +")'>Editar </button></li>";
                
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

  }


function teste(id){

                
              
                                    var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                      
                                    var sql = "UPDATE usuario SET UserName = '"+document.getElementById("firstName1Edit").value+"',LastName = '"+document.getElementById("lastName1Edit").value+"' WHERE id = 27 ";
                                  
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                    RS.close();


      
  //document.write("teste"+document.getElementById("firstName1Edit").value,document.getElementById("lastName1Edit").value);
}

  function Update(id,nome,subrenome){


    /*
			var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                    	
                                    var sql = "UPDATE usuario SET UserName = 'TESTE200',LastName = 'TESTE55' WHERE id = "+id;
                                  
                                    ConnectObj.Open (myConnect);
                                    RS.Open(sql,ConnectObj);
                                    RS.close();
                                    //alert("Registro alterado com sucesso!");
                                    //console.log("Registro alterado com sucesso!");
  */


var myConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\v_teste_map\\db_teste.mdb;Jet OLEDB:Database Password=123456";
                                    var ConnectObj = new ActiveXObject("ADODB.Connection");
                                    var RS = new ActiveXObject("ADODB.Recordset");
                                      //id,nome,subrenome
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


 //   document.onmousedown=protegercodigo

