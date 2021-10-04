import * as Vue from "vue"
import router from './router'
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
import _ from 'lodash'

import db from './database.js'

async function test(){
  try {
    await db.authenticate();
    console.log('Connection has been established successfully.');
  } catch (error) {
    console.error('Unable to connect to the database:', error);
  }
  
}

test()

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
