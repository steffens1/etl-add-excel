import Vue from 'vue'
import Router from 'vue-router'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/',
      name: 'home',
      component: require('../pages/Home').default
      
    },
    {
      path: '*',
      redirect: '/'
    }
  ]
})
