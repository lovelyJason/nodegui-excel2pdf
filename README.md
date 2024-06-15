nodegui-excel2pdf
======================
一个跨平台的，即 [NodeGui](https://docs.nodegui.org/) 项目,功能是把excel的表头和内容都按照word文档中的变量替换掉，一键生成word或pdf

<img width="360" alt="image" src="https://github.com/lovelyJason/nodegui-excel2pdf/assets/50656459/4c0f8429-8033-4701-862d-f7f99a6459c2">



安装
------
```bash
npm install
```

构建
--------
```bash
npm run build
```
使用esbuild进行构建

运行
-------
```bash
npm run start:watch
```

打包
---------
`npm run package` will run [Jam Pack NodeGui](https://github.com/sedwards2009/jam-pack-nodegui) with a configuration file to create the relevant packages for the current operating system this is running on. The output appears in `tmp-jam-pack-nodegui/jam-pack-nodegui-work/`.

或者使用

```bash
 npx nodegui-packer --pack ./dist
```


脚本：
------------------

* `build` - Runs all of the build steps.
* `build-code` - Just run in the TypeScript compiler.
* `build-bundle` - Run `esbuild` to create the output bundle file in `dist`.
* `clean` - Deletes the temporary files in `build` and `dist`.
* `run` - Runs the application from the `dist` folder.
* `package` - Build packages for the application. The output appears in `tmp-jam-pack-nodegui/jam-pack-nodegui-work/`
