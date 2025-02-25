<template>
  <span>Click the following button to get the property of content controls in a Word document</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "ContentControl.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Get all structureTags in the Word document
        let structureTags = GetAllTags(document);

        //Get all StructureDocumentTagInline objects
        let tagInlines = structureTags.tagInlines;

        let property = "";
        property += "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " + "\r\n";
        //Get properties of all tagInlines
        for (let i = 0; i < tagInlines.length; i++) {
            let alias = tagInlines[i].SDTProperties.Alias;
            let id = tagInlines[i].SDTProperties.Id;
            let tag = tagInlines[i].SDTProperties.Tag;
            let STDType = tagInlines[i].SDTProperties.SDTType.toString();
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
        }

        //Get all StructureDocumentTag objects
        let tags = structureTags.tags;

        //Get properties of all tags
        for (let i = 0; i < tags.length; i++) {
            let alias = tags[i].SDTProperties.Alias;
            let id = tags[i].SDTProperties.Id;
            let tag = tags[i].SDTProperties.Tag;
            let STDType = tags[i].SDTProperties.SDTType.toString();
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
        }
        
        // Define the output file name
        const outputFileName = "GetContentControlProperty.txt";

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, property);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

        // Clean up resources
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };
    const GetAllTags = (document) => {
        let tagInlines = [];
        let tags = [];

        let structureTags = new StructureTags(tagInlines, tags);
        for (let i = 0; i < document.Sections.Count; i++) {
            let section = document.Sections.get_Item(i);
            for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
                let obj = section.Body.ChildObjects.get(j);
                if (obj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
                    structureTags.tags.push(obj);
                } else if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
                    for (let j = 0; j < obj.ChildObjects.Count; j++) {
                        let pobj = obj.ChildObjects.get(j);
                        if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                            structureTags.tagInlines.push(pobj);
                        }
                    }
                }
                else if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Table) {
                    for (let a = 0; a < obj.Rows.Count; a++) {
                        let row = obj.Rows.get_Item(a);
                        for (let b = 0; b < row.Cells.Count; b++) {
                            let cell = row.Cells.get_Item(j);
                            for (let c = 0; c < cell.ChildObjects.Count; c++) {
                                let cellChild = cell.ChildObjects.get(c);
                                if (cellChild.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
                                    structureTags.tags.push(cellChild);
                                } else if (cellChild.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
                                    for (let d = 0; d < cellChild.ChildObjects.Count; d++) {
                                        let pobj = cellChild.ChildObjects.get(d);
                                        if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                                            structureTags.tagInlines.push(pobj);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        return structureTags;
    };
    
    class StructureTags {
        constructor(tagInlines, tags) {
            this.tagInlines = tagInlines;
            this.tags = tags;
        }
    }
    
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
