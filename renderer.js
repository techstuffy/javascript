//-------------------------------------------
// Renderer - create the browser window output
//-------------------------------------------
var interProcessComunicaton = require('electron').ipcRenderer;
$ = window.jQuery = require('jquery');

interProcessComunicaton.on('someMessage', function(event, argument,hostname){

	var html = '<table class="table table-striped">';
	var data = argument;
    html += '<thead><tr>';
    var flag = 0;	
	
    $.each(data[0], function(index, value){
		
//------------------------------------------------				
//Ignore the appname column header		
//------------------------------------------------				
		if(index != "appname"){
        html += '<th>'+index+'</th>';
		}		

    });
	 
    html += '</tr></thead><tbody>';
	
	
    $.each(data, function(index, value){
		
//------------------------------------------------		
//Check if this is a record at the current location		
//------------------------------------------------			
	if(value['location']==hostname){			
			
         html += '<tr>';
        $.each(value, function(index2, value2){
		
//------------------------------------------------				
//Ignore the appname column 		
//------------------------------------------------		
			if(index2 != "appname"){					
				html += '<td>'+value2+'</td>';			
			}
					
        });
        html += '<tr>';
		
	}
     });
	 
    html += '</tbody></table>';

	
     $('#mytable').html(html);
});
