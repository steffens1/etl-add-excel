<template>
  <div>
    <el-button @click="enviardata">enviar datta</el-button>
    <el-input v-model="slice"></el-input>
  </div>
</template>


<script>

  export default {
    data() {
      return {
        SheetList: [],
        SheetCategorized:[],
        data: {},
        slice :  100
      }
    },

    created(){
      Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;
        var arraySheets = sheets.load("items/name");
        var sheetPrecio = context.workbook.worksheets.getItem("L  Exclusiva C Porta  12M");
        var tablaRange = sheetPrecio.getUsedRange().load("values")
        await context.sync() 
        var headers = _.compact(tablaRange.values[0])
        var values = []
        tablaRange.values.map((item)=>{
         values.push(_.zipObject(headers, item ))
        })
        this.data = values     
      })
    },
    methods : {
      enviardata : function (){
        var sliceSize = parseInt(this.slice)
        const resul = this.data.slice(1,sliceSize)
        this.$http.post('/update-precios-productos', {
          data : resul
        })
      }
    }
  };
  
</script>

