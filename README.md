# PPT页生成插入音频器
使用醋栗_Stachelbeere大佬的PPT自动生成器从log.txt生成PPT:
https://www.bilibili.com/video/BV19B4y1A7nG?spm_id_from=333.788.videopod.episodes&vd_source=cf47911572dea905c360c43bf709f8b6&p=2


因为没有看到实装朗读女功能,所以做了一个简单的插入音频器,可以在PPT页插入音频文件,然后播放PPT就可以做成简易跑团视频了。exe执行文件在zip包中。


暂时使用阿里云接口（3个月试用）,需要自行申请阿里云ACCESS_KEY和APP_KEY，官网教程链接：
https://help.aliyun.com/zh/isi/getting-started/start-here?spm=a2c4g.11186623.help-menu-30413.d_1_0.720d6a85YtRsZs&scm=20140722.H_72138._.OR_help-T_cn~zh-V_1


音色excel里的内容可根据以下链接替换：
https://help.aliyun.com/zh/isi/developer-reference/overview-of-speech-synthesis?spm=a2c4g.11186623.0.0.43197cebK6WYph#5186fe1abb7ag


现有问题:
1. 阿里云接口只有3个月试用,等试用过期再看看有没有别的免费语音包方案,预计添加其他厂家语音合成API选择（讯飞开发中）。也可以先用朗读女生成全文音频，支持PPT自动切换和切分音频逐页插入（to be soon）
2. 插入的音频文件不能自动播放且有喇叭图标。（这个python-pptx库真有点。。。）

==========更新日志润色功能=============

textTool.py支持对log.txt中kp扮演npc的识别替换，使用阿里百炼模型tongyi-xiaomi-analysis-flash


==========更新朗读女支持===============

插入朗读女生成的全文音频，PPTAutoPlay.py根据lrc文件设置PPT自动切换
