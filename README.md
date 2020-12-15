> 在桌面标准化环境使用（如windows下），pip大于19.3版本，python大于3.7

安装模块，

```shell
Pillow，opencv-python，numpy，scikit-image
#需要使用合并ppt需要安装python-pptx模块
```

脚本具体帮助文档
```shell
python opencv.py -h
```
# 直播录屏

>可以使用脚本录屏或者专业的录屏软件
```shell
python opencv.py -r
```

# 提取图片

==视频里的动画变化跟人脸识别导致生成大量文件。==

- 均值哈希算法使用

```shel
python opencv.py -f file --ahash -s 0.98
```

默认-s相似度0.98。

- 只去除完全一样的图片，进行跳帧截屏时，类似间隔一段时间截图【不是很推荐】。

  ```shell
  python opencv.py -f file --same
  ```

- SSIM算法,使用的scikit-image模块,效果不错，推荐使用。【相比实现的ahash算法更慢，不过看上去更不错】

  相似信息多推荐使用0.98
  
  ```shell
python opencv.py -f file --ssim -s 0.98
  ```
  
  > ==增加跳帧功能，视频连贯性比较好时使用==，即舍弃一些画面。

默认跳帧10

```shell
python .\opencv.py -f .\video.avi --ahash -s 0.98 --frameskip 10
```

看连贯性，好的，可以考虑跳1000帧，具体看相似度（程序会打印出来）如何。

# 转ppt

挑选整理好图片后，使用程序直接转ppt，或者ppt中一次导入。

```shell
python opencv.py -p -d result
```



