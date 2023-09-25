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
                
                <!-- <input id="checkbox" type="checkbox"> -->
                <div class="inline-flex items-center">
                    <label
                        class="relative flex cursor-pointer items-center rounded-full p-3"
                        for="login"
                        data-ripple-dark="true"
                    >
                    <input
                    id="checkbox"
                    type="checkbox"
                    class="before:content[''] peer relative h-5 w-5 cursor-pointer appearance-none rounded-md border border-blue-gray-200 transition-all before:absolute before:top-2/4 before:left-2/4 before:block before:h-12 before:w-12 before:-translate-y-2/4 before:-translate-x-2/4 before:rounded-full before:bg-blue-gray-500 before:opacity-0 before:transition-opacity checked:border-green-500 checked:bg-green-500 checked:before:bg-green-500 hover:before:opacity-10 text-green-500 hover:bg-green-100"
                    />
                    <div class="pointer-events-none absolute top-2/4 left-2/4 -translate-y-2/4 -translate-x-2/4 text-white opacity-0 transition-opacity peer-checked:opacity-100">
                        <svg
                            xmlns="http://www.w3.org/2000/svg"
                            class="h-3.5 w-3.5"
                            viewBox="0 0 20 20"
                            fill="currentColor"
                            stroke="currentColor"
                            stroke-width="1"
                        >
                            <path
                            fill-rule="evenodd"
                            d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z"
                            clip-rule="evenodd"
                            ></path>
                        </svg>
                        </div>
                    </label>
                    <label
                        class="mt-px cursor-pointer select-none font-light text-gray-700"
                        for="login"
                    >
                        First Row as filename
                    </label>
                </div>
                <br>
                 <!--<button onclick="parseExcelFile2()">Generate document</button>-->
                <button type="button" onclick="parseExcelFile2()" class="py-3 px-4 inline-flex justify-center items-center gap-2 rounded-md border border-transparent font-semibold text-green-500 hover:bg-green-100 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all text-sm dark:focus:ring-offset-gray-800">
                Button
                </button>
                <script id="test">
                    </script>

                <script>
                    const docx = document.getElementById("docx");
                    const excel = document.getElementById("excel");
                    var index = 0;
                    const zip2 = new JSZip();

                    function parseExcelFile2(inputElement=excel) {
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
                                var result = ''
                                var docs = [];
                                
                                workbook.worksheets.forEach(function (sheet) {
                                    var names = [];
                                    var values = [];

                                    for(let j=0; j<sheet._rows[0]._cells.length; j++){
                                        names[j] = sheet._rows[0]._cells[j].value;
                                    }
                                    for(let i=1; i<sheet._rows.length; i++){
                                        values = [];
                                        for(let j=0; j<sheet._rows[i]._cells.length; j++){
                                            if(typeof sheet._rows[i]._cells[j].value === 'object' && sheet._rows[i]._cells[j].value !== null)
                                                values.push([names[j],sheet._rows[i]._cells[j].value.result])
                                            else
                                                values.push([names[j],sheet._rows[i]._cells[j].value])
                                               
                                        }
                                        
                                        values = Object.fromEntries(values);
                                        console.log(values);

                                        generate(values, false, i);
                                    } 
                                    generate(values, true);
                                });
                            });

                        };
                        reader.readAsArrayBuffer(file);
                    }

                    function generate(values, final, i=0) {
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
                            // Output the document using Data-URI
                            if(final == true){
                                zip2.generateAsync({type:"blob"}).then(function(content) {
                                        saveAs(content, "example.zip");
                                    });
                            }
                            else{
                                const entries = Object.entries(values);
                                console.log(entries)
                                if(document.getElementById('checkbox').checked)
                                    zip2.file(entries[0][1]+".docx",  blob);
                                else
                                    zip2.file("output"+i+".docx",  blob);
                            }
                            //saveAs(blob, "output.docx");
                        };

                    }

                </script>
                </div>
            </div>
        </div>
    </div>
</x-app-layout>
