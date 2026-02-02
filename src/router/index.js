import { createRouter, createWebHashHistory } from 'vue-router'
//import HomeView from '../views/HomeView.vue'

const router = createRouter({
  history:  createWebHashHistory(''),
  routes: [
    {
      path: '/',
      name: '默认页',
      component: () => import('../components/Root.vue')
    },
    {
      path: '/dialog',
      name: '对话框',
      component: () => import('../components/Dialog.vue')
    },
    {
      path: '/mydialog',
      name: '对话框',
      component: () => import('../components/MyDialog.vue')
    },
    {
      path: '/taskpane',
      name: '任务窗格',
      component: () => import('../components/TaskPane.vue')
    },
    {
      path: '/materialquery',
      name: '物料查询',
      component: () => import('../components/MaterialQueryTaskPane.vue')
    },
    {
      path: '/calculator',
      name: '科学计算器',
      component: () => import('../components/ScientificCalculatorTaskPane.vue')
    },
    {
      path: '/currency',
      name: '汇率换算',
      component: () => import('../components/CurrencyConverterTaskPane.vue')
    }
  ]
})

export default router
