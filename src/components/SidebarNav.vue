<script setup lang="ts">
import { useRoute } from 'vue-router'

export interface NavChild {
  label: string
  route: string
}

export interface NavItem {
  label: string
  route?: string
  children?: NavChild[]
  sep?: boolean
}

withDefaults(
  defineProps<{
    navItems: NavItem[]
    activeDropdown?: number
  }>(),
  {
    activeDropdown: undefined,
  },
)

const emit = defineEmits<{
  (e: 'toggle-dropdown', idx: number): void
  (e: 'navigate', route: string): void
  (e: 'close-all'): void
}>()

const route = useRoute()
</script>

<template>
  <nav id="sidebar">
    <template v-for="(item, idx) in navItems" :key="idx">
      <div v-if="!item.children" class="nav-item">
        <div
          class="nav-label" :class="{ 'active-page': route.name === item.route }"
          @click="emit('navigate', item.route ?? '/'); emit('close-all')"
        >
          {{ item.label }}
        </div>
      </div>

      <div v-else class="nav-item">
        <div
          class="nav-label" :class="{ 'active-dropdown': activeDropdown === idx }"
          @click.stop="emit('toggle-dropdown', idx)"
        >
          {{ item.label }}
          <span class="chevron">▼</span>
        </div>

        <div v-if="activeDropdown === idx" class="dropdown-panel">
          <div class="dropdown-header">
            {{ item.label }}
          </div>
          <div
            v-for="child in item.children" :key="child.route" class="dropdown-item"
            :class="{ 'active-child': route.name === child.route }"
            @click="emit('navigate', child.route); emit('close-all')"
          >
            {{ child.label }}
          </div>
        </div>
      </div>

      <div v-if="item.sep" class="nav-separator" />
    </template>
  </nav>
</template>

<style scoped>
#sidebar {
  width: var(--sidebar-width);
  min-height: 100vh;
  background: var(--bg-sidebar);
  border-left: 1px solid var(--bg-sidebar-border);
  display: flex;
  flex-direction: column;
  padding: 24px 0;
  gap: 2px;
  position: relative;
  z-index: 10;
  flex-shrink: 0;
}

.nav-item {
  position: relative;
  width: 100%;
}

.nav-label {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 12px 16px;
  font-family: var(--font-nav);
  font-size: 15px;
  font-weight: 600;
  letter-spacing: 0.04em;
  color: var(--text-primary);
  cursor: pointer;
  border-radius: 6px;
  margin: 0 8px;
  transition: background 0.15s, color 0.15s;
  user-select: none;
  text-transform: uppercase;
}

.nav-label:hover {
  background: rgba(255, 255, 255, 0.07);
}

.nav-label.active-page {
  color: var(--accent);
}

.nav-label.active-dropdown {
  color: var(--accent);
}

.chevron {
  font-size: 10px;
  transition: transform 0.25s;
  color: var(--text-muted);
  margin-left: auto;
}

.active-dropdown .chevron {
  transform: rotate(180deg);
  color: var(--accent);
}

.dropdown-panel {
  position: relative;
  right: 0;
  top: 0;

  width: auto;
  margin: 4px 8px;

  background: var(--accent);
  border-radius: 15px;
  overflow: hidden;
  box-shadow: 0 8px 32px #ff202059, 0 2px 8px #00000066;
  z-index: 100;

  transform-origin: top center;
  animation: dropIn 0.20s cubic-bezier(0.35, 1.55, 0.65, 1);
}

@keyframes dropIn {
  from {
    opacity: 0;
    transform: scale(0.92) translateY(-6px);
  }

  to {
    opacity: 1;
    transform: scale(1) translateY(0);
  }
}

.dropdown-header {
  font-family: var(--font-nav);
  font-size: 15px;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-primary);
  padding: 14px 16px 10px;
  text-align: center;
}

.dropdown-item {
  display: block;
  text-align: center;
  padding: 11px 16px;
  font-family: var(--font-nav);
  font-size: 14px;
  font-weight: 600;
  letter-spacing: 0.03em;
  color: var(--text-primary);
  cursor: pointer;
  transition: background 0.12s;
  text-transform: uppercase;
}

.dropdown-item:hover {
  background: rgba(255, 255, 255, 0.18);
}

.dropdown-item:last-child {
  padding-bottom: 14px;
}

.dropdown-item.active-child {
  color: var(--accent);
  background: rgba(255, 255, 255, 0.25);
}

.nav-separator {
  height: 1px;
  background: var(--bg-sidebar-border);
  margin: 6px 16px;
}
</style>
