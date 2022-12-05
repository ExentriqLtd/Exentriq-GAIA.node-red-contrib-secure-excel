const XlsxGenerator = require("xlsx-encrypter");

module.exports = function(RED) {
	
    function SecureExcel(config) {
        RED.nodes.createNode(this,config);
        var node = this;
        node.on('input', function(msg) {
          
            const password = msg.pwd;
            const xlsxGenerator = new XlsxGenerator.XlsxGenerator();
            /*const data = [
			  {
			    fruit: "Apples2",
			    quantity: 4,
			    price: "£6.86",
			  },
			];*/
			
			const data = msg.payload;
			
			const sheetName = msg.sheetName || "Sheet 1";
			
			// Create worksheets
			const workSheet = xlsxGenerator.createWorkSheet(data, sheetName);
			
			// Create XLSX file
			if(password){
				void xlsxGenerator.exportWorkSheetsToFile(msg.filepath, [
				  workSheet
				],password);
			}else{
				void xlsxGenerator.exportWorkSheetsToFile(msg.filepath, [
				  workSheet
				]);
			}
			
            return msg;
            
        });
    }
    
    RED.nodes.registerType("secure-excel",SecureExcel);
}