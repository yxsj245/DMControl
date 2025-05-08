# DMControl
用于Windows设备管理器中指定USB硬件的快速禁用和启用
![image](https://github.com/user-attachments/assets/d0841540-7127-4cc7-9458-9c529522b6f9)

# 如何使用
从右侧[Releases](https://github.com/yxsj245/DMControl/releases)中下载最新发行二进制可执行程序
程序运行后会在运行目录下创建一个config.json 配置文件
```json
{
    "device_id": "USB\\VID_174C&PID_1153",
    "use_full_id": false,
    "full_device_id": "USB\\VID_174C&PID_1153\\MSFT3023456789013B"
}
```
前往 设备管理器 找到需要管理的硬件ID，比如硬盘盒一般是在存储控制器中 \
`device_id` ——硬件ID \
![](https://pic1.imgdb.cn/item/681c829558cb8da5c8e5ccc7.png) \
`use_full_id` 序是否使用完整设备ID(true)还是仅使用部分ID(false)。当设置为true时，程序会使用full_device_id字段中的值进行更精确的设备匹配 \
`full_device_id` ——设备实例路径(设备ID)\
![](https://pic1.imgdb.cn/item/681c838f58cb8da5c8e5ce0c.png) \
按照上述要求进行修改
### 记住一定要使用双斜杠！！
随后重新启动程序即可
