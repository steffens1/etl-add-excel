import * as Vue from "vue"
import router from './router'
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
import _ from 'lodash'
import "./assets/index.css"

import axios from './axios.js'


Vue.prototype.$http = axios

Vue.use(ElementUI)
window._ = _

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    new Vue({
      el: "#app",
      router
    })
  }
})
