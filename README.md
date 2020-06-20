# tranExcelUrlToPicture
读取Excel中的Url地址，并将其转换为图片
一、用Python实现
实现思路：
1、打开文件，选择需要转换的Excel
2、读取数据到dataframe
3、定位url所在列
4、读取所有url地址
5、遍历所有url地址，逐个将图片下载到本地
6、创建新Excel，将源表中所有内容插入，同时在Url所在列后面增加一列“插入图片”，插入图片

二、用Excel宏实现
1、新建Excel，界面打开后，依次点击 文件-选项-自定义功能区，主选项卡中勾选“开发工具”
2、保存，保存类型选择“Excel启用宏的工作薄（*.xlsm)",录入文件名
3、依次点击 开发工具-查看代码，会弹出的VBA设计界面
4、双击 Microsoft Excel对象下的sheet，默认为 Sheet1(Sheet1)
5、粘贴VBA代码，然后关闭。
6、依次点击 开发工具-宏，选择 “Sheet1.UrlDownloadPicture”，点击执行。
7、依次录入提示信息，即可完成转换。
