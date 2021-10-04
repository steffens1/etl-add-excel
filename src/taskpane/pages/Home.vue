<template>
  <el-table
    :data="data"
    stripe
    style="width: 100%">
    <el-table-column
      prop="NOMBRE COMERCIAL"
      label="NOMBRE COMERCIAL">
    </el-table-column>
    
    <el-table-column
      prop="PLAN"
      label="PLAN">
    </el-table-column>
    
    <el-table-column
      prop="Valor"
      label="Valor">
    </el-table-column>
    
    <el-table-column
      prop="MODALIDAD"
      label="MODALIDAD">
    </el-table-column>
    
    <el-table-column
      prop="CAMPANA"
      label="CAMPANA">
    </el-table-column>

  </el-table>
</template>


<script>
  export default {
    data() {
      return {
        SheetList: [],
        SheetCategorized:[],
        data: {}
      }
    },
    created(){
      Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;
        var arraySheets = sheets.load("items/name");

        var sheetPrecio = context.workbook.worksheets.getItem("CUOTAS RENO 12M");
        var tablaRange = sheetPrecio.getUsedRange().load("values")

        
        await context.sync() 

        var headers = _.compact(tablaRange.values[0])
        var values = []

        tablaRange.values.map((item)=>{
         values.push(_.zipObject(headers, item ))
        })

        this.data = values
        console.log(values)
      
      })
    }
  };
</script>

