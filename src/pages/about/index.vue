<script setup lang="ts">

</script>

<template>
  <div>
    <div class="page-subtitle">
      Proje Hakkında
    </div>
    <div class="page-title">
      <a href="https://github.com/iseeeva/uyap" target="_blank" rel="noopener noreferrer">iseeeva/uyap</a>
    </div>
    <div class="divider" />
    <div class="page-body">
      <p>
        Bu proje, <strong>UYAP Editör</strong> ile oluşturulan formatlar üzerinde kolayca işlem yapabilmeniz için
        geliştirilmiş <strong>açık kaynaklı bir toolbox'dır</strong>.
      </p>
      <p class="disclaimer">
        <strong>Bilgilendirme:</strong> Bu projenin UYAP (Ulusal Yargı Ağı Bilişim Sistemi), Adalet Bakanlığı veya
        resmi kurum ve iştirakleriyle hiçbir <strong>bağı, ortaklığı veya resmi sorumluluğu bulunmamaktadır</strong>.
      </p>
    </div>
  </div>
</template>

<style scoped>
a {
  color: var(--accent);
  background-color: transparent;
  text-decoration: none;
  word-break: break-word;
}

a:hover,
a:active {
  text-decoration: underline;
}

.page-body {
  max-width: min(720px, 90vw);
}

.page-body p {
  line-height: 1.6;
  margin-bottom: 1.5rem;
  font-size: clamp(14px, 1.5vw + 10px, 17px);
}

.disclaimer {
  background-color: rgba(202, 62, 71, 0.12);
  border-left: 4px solid var(--accent);
  border-radius: 4px;
  padding: clamp(10px, 1.5vw, 16px) clamp(12px, 2vw, 18px);
  color: var(--text-primary);
  line-height: 1.5;
}
</style>
