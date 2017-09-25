# RobotHostSoftware  

# 直角坐标机器人控制上位机

## 使用winform实现的直角坐标机器人的上位机
  
**注：最终有一个三维坐标并显示该坐标点的肌张力的数值的功能尚未实现**
  
### 功能说明  
     
    ● 实时接收来自机器人控制器通过串口发送上来的的机器人x轴、Y轴、Z轴以及力传感器轴的实时位置信息；  
    
    ● 利用4个chart控件上进行相关数据的实时动态显示；    
    
    ● 可接收到的各轴位置信息以及力传感器的值以excel表格的形式进行存储，用NPOI实现；  
    
    ● 发送指令，可分为手动模式以及自动模式。手动模式是点击各轴相应按钮使各轴运动；自动控制模式是给定空间中的坐标进行给定；  
    
    ● 包含串口通行的协议；  
    
    ● 对历史所存的数据（Excel格式）可以利用此软件直接打开并进行可视化处理；  
    
    
 
 ## 界面截图
![image](https://github.com/ligb1023561601/RobotHostSoftware/raw/master/screenshot/捕获.PNG)

![image](https://github.com/ligb1023561601/RobotHostSoftware/raw/master/screenshot/捕获1.PNG)

