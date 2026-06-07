<script setup lang="ts">
import type { NavItem } from './components/SidebarNav.vue'
import { ref } from 'vue'
import { useRouter } from 'vue-router'
import SidebarNav from './components/SidebarNav.vue'

const router = useRouter()

const navItems: NavItem[] = [
  { label: 'udf', route: 'udf' },
  // {
  //   label: 'Dropdown 1',
  //   children: [
  //     { label: 'Item 1.1', route: 'item-1-1' },
  //     { label: 'Item 1.2', route: 'item-1-2' },
  //     { label: 'Item 1.3', route: 'item-1-3' },
  //   ],
  // },
  { label: 'hakkinda', route: 'about' },
]

const activeDropdown = ref<number | undefined>(undefined)

function toggleDropdown(idx: number): void {
  activeDropdown.value = activeDropdown.value === idx ? undefined : idx
}

function closeAll(): void {
  activeDropdown.value = undefined
}

function navigate(routeName: string): void {
  router.push({ name: routeName })
}
</script>

<template>
  <div id="app-wrapper">
    <div v-if="activeDropdown != null" class="overlay" @click="closeAll" />

    <main id="main-content">
      <router-view v-slot="{ Component }">
        <transition name="fade" mode="out-in">
          <component :is="Component" />
        </transition>
      </router-view>
    </main>

    <SidebarNav
      :nav-items="navItems" :active-dropdown="activeDropdown" @toggle-dropdown="toggleDropdown"
      @navigate="navigate" @close-all="closeAll"
    />
  </div>
</template>
