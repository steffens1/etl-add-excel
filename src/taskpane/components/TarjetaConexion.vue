<template>
  <div class="">
    <el-card class="w-full">
      <el-form label-width="120px" class="hoja--formulario">
        <el-form-item class="">
          <label for="">Endpoint</label>
          <el-input v-model="hoja.endpoint"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Coleccion</label>
          <el-input v-model="hoja.coleccion"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Hoja</label>
          <el-input v-model="hoja.hoja"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Conexion</label>
          <el-input v-model="hoja.conexion"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Grupo</label>
          <el-input v-model="hoja.grupo"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Nombre</label>
          <el-input v-model="hoja.nombre"></el-input>
        </el-form-item>
        <el-form-item>
          <label for="">Campos</label>
          <div class="flex justify-between space-x-2">
            <draggable class="w-1/2 space-y-2" v-model="local" group="interno">
              <div v-for="(item, index) in local" :key="`i-${index}`">{{item}}</div>
            </draggable>
            <draggable class="w-1/2 space-y-2" v-model="servidor" group="externo">
              <div
                v-for="(item, index) in servidor"
                :key="`s-${index}`"
                class="rounded shadow-lg text-blue-400 text-right pr-3"
              >
                {{item}}
              </div>
            </draggable>
          </div>
        </el-form-item>
      </el-form>
    </el-card>
  </div>
</template>

<script>
import draggable from 'vuedraggable'
export default {
  components: {
    draggable
  },
  props:{
    hoja: Object,
  },
  data() {
    return {
      local: [0,1,2,3],
      servidor: ['a', 'b', 'c', 'd']
    }
  },
  methods: {
  }
}
</script>

<style>
.hoja--formulario .el-form-item__content{
  margin-left: 8px !important;
}
</style>