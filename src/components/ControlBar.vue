<!-- SPDX-License-Identifier: MIT -->
<script lang="ts">
import VueFileToolbarMenu from "vue-file-toolbar-menu";

let tbDataForJs;

export default {
  components: { VueFileToolbarMenu },
  data() {
    return {
      active: {
        Bold: false,
        Italic: false,
        Underline: false
      },
      disabled: true,  // startup default
    };
  },

  computed: {
    my_menu() {
      tbDataForJs = this;  // line may run multiple times
      return [
        {
          icon: "format_bold", title: "Bold",
          click: () => toggleFormatting('Bold'),
          active: this.active['Bold'],
          disabled: this.disabled,
        },
        {
          icon: "format_italic", title: "Italic",
          click: () => toggleFormatting('Italic'),
          active: this.active['Italic'],
          disabled: this.disabled,
        },
        {
          icon: "format_underline", title: "Underline",
          click: () => toggleFormatting('Underline'),
          active: this.active['Underline'],
          disabled: this.disabled,
        },
      ];
    },
  },
  init_control_bar: function() {
    jsPassCtrlBar(tbDataForJs);
  },
};
</script>




<template>
  <div>
    <vue-file-toolbar-menu :content="my_menu" />
  </div>
</template>

<style scoped>
</style>
