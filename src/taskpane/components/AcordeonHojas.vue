<template>
  <div class="">
    <div v-for="(item, index) in tipos" :key="`tipo-${index}`">
      <span>{{item}}</span>
      <div class="p-4">
        <el-collapse v-model="hojaActiva" accordion>
          <el-collapse-item v-for="(hoja, indexh) in hojas" :key="`hoja-${indexh}`" :name="indexh">
            <template v-slot:title>
              <div class="w-full flex">
                <span class="item-formulario--titulo">{{hoja.nombreHoja}}</span>
                <div class="mr-2">
                <el-button icon="el-icon-upload" circle size="small"/>
                <el-button icon="el-icon-setting" circle size="small"/>
                <el-button icon="el-icon-refresh" circle size="small"/>
                </div>
              </div>
            </template>
            <tarjeta-conexion :hoja="hoja"/>
          </el-collapse-item>
        </el-collapse>
      </div>
    </div>
  </div>
</template>

<script>
import TarjetaConexion from './TarjetaConexion.vue'
export default {
  props: {
    tipos: Array,
    hojas: Array
  },
  components: {
    TarjetaConexion
  },
  data() {
    return {
      hojaActiva: null
    }
  }
}
</script>

<style>
.item-formulario--titulo {
  margin-right: auto !important;
}
</style>