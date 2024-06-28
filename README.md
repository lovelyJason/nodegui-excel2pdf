nodegui-excel2pdf
======================
一个跨平台的，即 [NodeGui](https://docs.nodegui.org/) 项目,功能是把excel的表头和内容都按照word文档中的变量替换掉，一键生成word或pdf

<img width="360" alt="image" src="https://github.com/lovelyJason/nodegui-excel2pdf/assets/50656459/4c0f8429-8033-4701-862d-f7f99a6459c2">

这个项目是个好项目（我指的是nodegui这个框架），只是社区不再维护了， 好好的东西成了一坨

其他跨平台对比
------

Tauri:

<img width="1316" alt="image" src="https://github.com/lovelyJason/nodegui-excel2pdf/assets/50656459/2456c4f8-a23a-483d-ac14-370478ff789a">

flutter:

<img width="1276" alt="image" src="https://github.com/lovelyJason/nodegui-excel2pdf/assets/50656459/afc02f53-3efd-4c45-a115-4f3e695f6819">

没错，flutter现在也支持了桌面应用，包括windows和macos， 还有什么理由不选择呢, nodegui这套风格我还是喜欢的， 对于不想去学pyqt的还是有点用，奈何社区不给力，好多bug，也不维护了

这个需求用flutter实现了， 项目地址[https://github.com/lovelyJason/flutter_excel2pdf](https://github.com/lovelyJason/flutter_excel2pdf)

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
