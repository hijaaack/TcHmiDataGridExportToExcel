// Keep these lines for a best effort IntelliSense of Visual Studio 2017 and higher.
/// <reference path="./../../Packages/Beckhoff.TwinCAT.HMI.Framework.12.758.8/runtimes/native1.12-tchmi/TcHmi.d.ts" />

(async function (/** @type {globalThis.TcHmi} */ TcHmi) {
    // If you want to unregister an event outside the event code you need to use the return value of the method register()
    let destroyOnInitialized = TcHmi.EventProvider.register('TcHmiHtmlHost.onAttached', (e, data) => {
        // This event will be raised only once, so we can free resources. 
        // It's best practice to use destroy function of the event object within the callback function to avoid conflicts.
        e.destroy();

        //Add event for file-input button
        document.getElementById('file-input').addEventListener('change', readSingleFile, false);

        async function readSingleFile(e) {
            var file = e.target.files[0];
            readImportData(file);
        }

        async function readImportData(file) {

            const wb = new ExcelJS.Workbook();
            const reader = new FileReader()

            let newArr = [];
            let obj = {};

            reader.readAsArrayBuffer(file)
            reader.onload = () => {
                const buffer = reader.result;
                wb.xlsx.load(buffer).then(workbook => {
                    //console.log(workbook, 'workbook instance')
                    workbook.eachSheet((sheet, id) => {
                        sheet.eachRow((row, rowIndex) => {
                            //Skip first index, no need to read the excel headers, you could use the headers to dynamic create an object
                            if (rowIndex > 1) {
                                obj = {
                                    "ProductName": row.values[1],
                                    "Quality": row.values[2],
                                    "Value": row.values[3],
                                    "Verified": row.values[4]
                                };
                                newArr.push(obj);
                            }
                            //console.log(rowIndex);
                            //console.log(row.values, rowIndex)
                        })
                    })
                    //Set imported data to datagrid
                    const ctrl = TcHmi.Controls.get("TcHmiDatagrid_1");
                    ctrl.setSrcData(newArr);
                });
              
            }
        }

    });
})(TcHmi);
