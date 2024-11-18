Mentor Xpedition Automation

Script的加载
    1、使用绝对路径加载Script,如[run c:\temp\ViaCount.vbs]
    2、Application主动搜索路径加载Script，如[run ViaCount.vbs],该方法需要将Script放在如下路径下
        1）Product Area 产品安装目录下，由SDD_HOME环境变量指定的路径如：%SDD_HOME%\standard\automation\startup
        2）Design Area 设计文件目录下，包含*.pcb的目录
        3）User-Defined Areas 由WDIR环境变量指定的目录
    WDIR环境变量可以指定多个目录，使用（;）分割开

Script的应用启动加载和执行
    1、使用script.ini文件，当应用启动时会自动加载script.ini中指定的脚本文件
    2、script.ini文件可以放在Product Area, Design Area, User-Defined Area，且可以存在多个script.ini文件，应用会逐一搜寻和处理script.ini文件
    3、script.ini文件内容，以[Product name]起始指定该脚本运行在哪个应用中，脚本文件必须以Script#0开始，并且不能打乱或间断数字顺序，script.ini只执行从起始开始顺序排序的脚本
        [Expedition PCB]
        Script#0=pcbeevm.vbs
        Script#1=ShadowMovePart.vbs
        Script#2=Activelayerkeybindings.vbs
        Script#3=meizuPcbmenu.vbs

        [Xpedition PD]
        Script#0=pcbeevm.vbs
        Script#1=ShadowMovePart.vbs
    4、若需要打开任何设计时加载并自动执行脚本，在[Product name]后附加"- Document"限定符
        [Expedition PCB]
        Script#0=C:\temp\ViaCountMenuSingle.vbs
        [Expedition PCB - Document]
        Script#0=C:\temp\ViaCount.vbs
        则当打开应用时会自动执行ViaCount.vbs
