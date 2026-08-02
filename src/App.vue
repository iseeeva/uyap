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

<style scoped>
#app-wrapper {
  display: flex;
  height: 100dvh;
  width: 100%;
  overflow: hidden;
}

#main-content {
  flex: 1;
  min-width: 0;
  height: 100dvh;
  padding: 25px 40px;
  overflow-y: auto;
  overscroll-behavior: contain;
}

.overlay {
  position: fixed;
  inset: 0;
  z-index: 9;
  -webkit-tap-highlight-color: transparent;
}

.fade-enter-active,
.fade-leave-active {
  transition: opacity 0.18s ease, transform 0.18s ease;
}

.fade-enter-from {
  opacity: 0;
  transform: translateX(-10px);
}

.fade-leave-to {
  opacity: 0;
  transform: translateX(10px);
}

/* ---------- Phone ---------- */
@media (max-width: 640px) {
  #app-wrapper {
    flex-direction: column;
  }

  #main-content {
    height: 100dvh;
    padding: 18px 16px;
    padding-left: max(16px, env(safe-area-inset-left));
    padding-right: max(16px, env(safe-area-inset-right));
    padding-bottom: calc(var(--bottom-nav-height, 64px) + env(safe-area-inset-bottom) + 18px);
  }
}
</style>
