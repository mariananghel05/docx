<x-app-layout>
    <x-slot name="header">
        <h2 class="font-semibold text-xl text-gray-800 dark:text-gray-200 leading-tight">
            {{ __('Dashboard') }}
        </h2>
    </x-slot>

    <div class="py-12">
        <div class="max-w-7xl mx-auto sm:px-6 lg:px-8">
            <div class="bg-white dark:bg-gray-800 overflow-hidden shadow-sm sm:rounded-lg">
                <div class="p-6 text-gray-900 dark:text-gray-100">
                <script src="./js/docxtemplater.js"></script>
                <script src="./js/pizzip.js"></script>
                <script src="./js/FileSaver.js"></script>
                <script src="./js/pizzip-utils.js"></script>
                <script src="./js/exceljs.min.js"></script>
                <script src="https://cdn.jsdelivr.net/npm/docx-merger@1.2.2/dist/docx-merger.min.js"></script>
                <script src="./js/jszip.js" ></script>

                <!--<input type="file" id="docx" class="" />-->
                <h1>Choose .docx file</h1>
                <form>
                    <label for="docx" class="sr-only">Choose .docx file</label>
                    <input type="file" name="docx" id="docx" class="block w-full border border-gray-200 shadow-sm rounded-md text-sm focus:z-10 focus:border-blue-500 focus:ring-blue-500 dark:bg-slate-900 dark:border-gray-700 dark:text-gray-400
                        file:bg-transparent file:border-0
                        file:bg-gray-100 file:mr-4
                        file:py-3 file:px-4
                        dark:file:bg-gray-700 dark:file:text-gray-400">
                </form>
                <h1>Choose .excel file</h1>
                <!-- <input type="file"  id="excel"> -->
                <form>
                    <label for="file-input" class="sr-only">Choose .excel file</label>
                    <input type="file" name="file-input" id="excel" class="block w-full border border-gray-200 shadow-sm rounded-md text-sm focus:z-10 focus:border-blue-500 focus:ring-blue-500 dark:bg-slate-900 dark:border-gray-700 dark:text-gray-400
                        file:bg-transparent file:border-0
                        file:bg-gray-100 file:mr-4
                        file:py-3 file:px-4
                        dark:file:bg-gray-700 dark:file:text-gray-400">
                </form>


                <button type="button" onclick="start()" class="py-3 px-4 inline-flex justify-center items-center gap-2 rounded-md border border-transparent font-semibold text-green-500 hover:bg-green-100 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all text-sm dark:focus:ring-offset-gray-800">
                Button
                </button>

                <script id="test">
                    </script>

                <script>
                    const docx = document.getElementById("docx");
                    const excel = document.getElementById("excel");
                    var merging_contents = [];
                    

                    function read(file, final){
                        const reader = new FileReader();
                        reader.readAsBinaryString(file);
                        reader.addEventListener('load', function (evt) {
                            const content = evt.target.result;
                            merging_contents.push(content);
                            if(final)
                                merge_docx(merging_contents);
                        })
                    }

                    function generate(values, i=0, final) {
                        const reader = new FileReader();
                    
                        if (docx.files.length === 0) {
                            alert("No files selected");
                        } 
                        reader.readAsBinaryString(docx.files.item(0));

                        reader.onerror = function (evt) {
                            console.log("error reading file", evt);
                            alert("error reading file" + evt);
                        };
                        reader.onload = function (evt) {
                            const content = evt.target.result;
                            
                            const zip = new PizZip(content);
                            const doc = new window.docxtemplater(zip, {
                                paragraphLoop: true,
                                linebreaks: true,
                            });

                            // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
                            doc.render(values);
                            
                            const blob = doc.getZip().generate({
                                type: "blob",
                                mimeType:
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                // compression: DEFLATE adds a compression step.
                                // For a 50MB output document, expect 500ms additional CPU time
                                compression: "DEFLATE",
                            });

                            read(new File([blob], "output"+i), final)
                        };

                    }
                    function merge_docx (array)  {
                        var docx = new DocxMerger({},array);

                        docx.save('blob',function (data) {
                            saveAs(data,"output.docx");
                        });
                    }
                    function makefiles(objects){
                        for(var i=0; i<objects.length;i++){
                            if(i==objects.length-1)
                                generate(objects[i],i,true)
                            else
                                generate(objects[i],i,false)
                        }
                        console.log(merging_contents)
                        
                    }

                    function parseExcelFile2(inputElement=excel) {
                        var objects = [];
                        var files = inputElement.files || [];
                        if (!files.length) return;
                        var file = files[0];

                      

                        console.time();
                      
                        var reader = new FileReader();
                        reader.onloadend = function(event) {
                            var arrayBuffer = reader.result;
                            var workbook = new ExcelJS.Workbook();

                            // workbook.xlsx.read(buffer)
                            workbook.xlsx.load(arrayBuffer).then(function(workbook) {
                                console.timeEnd();
                              
                                workbook.worksheets.forEach(function (sheet) {
                                    var names = [];
                                    var values = [];
                                

                                    for(let j=0; j<sheet._rows[0]._cells.length; j++){
                                        names.push(sheet._rows[0]._cells[j].value)
                                    }
                                    for(let i=1; i<sheet._rows.length; i++){
                                        values = [];
                                        for(let j=0; j<sheet._rows[i]._cells.length; j++){
                                            values.push([names[j], sheet._rows[i]._cells[j].value])
                                        }
                                    objects.push(Object.fromEntries(values))  
                                  }
                                  makefiles(objects)
                              });
                          });

                      };
                      reader.readAsArrayBuffer(file);
                    }
                    function start(){
                        parseExcelFile2(excel);
                    }
                </script>
                </div>
            </div>
        </div>
    </div>
</x-app-layout>
