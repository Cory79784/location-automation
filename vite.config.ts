import {defineConfig} from 'vitest/config';
export default defineConfig({base:'./',build:{target:'es2022',rollupOptions:{output:{inlineDynamicImports:true}}},test:{pool:'forks',maxWorkers:2,minWorkers:1}});
