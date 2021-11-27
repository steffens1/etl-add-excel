<template>
  <div>
    <acordeon-hojas :tipos="tiposHoja" :hojas="hojas"/>
  </div>
</template>


<script>
import TarjetaConexion from '../components/TarjetaConexion.vue'
import AcordeonHojas from '../components/AcordeonHojas.vue'
  export default {
    components: {
      TarjetaConexion,
      AcordeonHojas
    },
    data() {
      return {
        SheetList: [],
        SheetCategorized:[],
        data: {},
        hojas: [],
        tiposHoja: []
      }
    },

    created(){
      console.log(Excel.QueryCollection)
      Excel.run(async (context) => {
        // var sheets = context.workbook.worksheets;
        // var arraySheets = sheets.load("items/name");

        // var sheetPrecio = context.workbook.worksheets.getItem("CUOTAS RENO 12M");
        // var tablaRange = sheetPrecio.getUsedRange().load("values")

        
        // await context.sync() 

        // var headers = _.compact(tablaRange.values[0])
        // var values = []

        // tablaRange.values.map((item)=>{
        //  values.push(_.zipObject(headers, item ))
        // })

        // this.data = values
        // console.log(values)
        let config = context.workbook.worksheets.getItem('Settings')
        let tabla = config.getUsedRange().load('values')
        await context.sync()
        let cabezeras = tabla.values[0]
        console.log('cabezeras', cabezeras)
        const tiposDeHojas = new Set()
        for( let i = 1 ; i < tabla.values.length ; i++){
          let temp = {}
          for( const [index, item] of tabla.values[i].entries()){
            temp[cabezeras[index]] = item
          }
          this.hojas.push(temp)
          console.log(temp.nombreHoja)
          tiposDeHojas.add(temp.tipo)
        }
        this.tiposHoja = Array.from(tiposDeHojas)
      })
    },
    methods: {}
  };
  
</script>

