import { createRouter, createWebHashHistory } from 'vue-router'

import AboutPage from '../pages/about/index.vue'
import UdfPage from '../pages/udf/index.vue'

export default createRouter({
  history: createWebHashHistory(),
  routes: [
    { path: '/', name: '', component: AboutPage },
    { path: '/about', name: 'about', component: AboutPage },
    { path: '/udf', name: 'udf', component: UdfPage },
    { path: '/:pathMatch(.*)*', redirect: '/' },
  ],
})
