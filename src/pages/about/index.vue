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
}

a:hover,
a:active {
  text-decoration: underline;
}

.page-body p {
  line-height: 1.6;
  margin-bottom: 1.5rem;
}

.disclaimer {
  background-color: #fff9e6;
  border-left: 4px solid #ffcc00;
  border-radius: 4px;
  padding: 12px 16px;
  color: #5c4b00;
  line-height: 1.5;
  font-size: 0.95rem;
}
</style>
