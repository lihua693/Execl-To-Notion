# Notion.py使用说明

#### notion.py --[options]

###### options:

* -h / --help 使用说明
* -c / --config 配置文件
* -d / --finish_path 指定数据添加完后备份的目标目录
* -s / --source_path 指定数据目录



## 配置文档说明

##### 必选参数

* source : --source_path 数据目录
* finish ： -d / --finish_path 数据添加完毕后的目标目录
* token ： 浏览器中cookies中的token_v2信息
* url ： notion中需要获取的页面url （必须是一级目录）
* floor ： 插入表的层级关系 如：["第二层title","第三层title","table title ‘kmos 2’ "]代表目标表在第四层,表名为kmos 2
* primeKey ： 主键
* info ： execl中与notion表对应的信息（非常重要，请认真填写）
* primeKey ： 设置主键，值为execl中其中一个表头

```
"info": [
  {
    "Execl": "execl中的title",
    "Notion": "notion表中的title",
    "Type": "数据类型" //目前仅支持str， date， int
  }
  ]
```

#### 注意事项：

* notion表中如果是标签类型或者select类型，必须首先手动创建一个标签或者select ！！！
* info中所对应的信息必须一一对应，若execl中没有，notion中有则会忽略添加，若notion中没有，execl中有则会报错。
* 数据类型目前仅支持 str， date， int

#### config 文件 example:

```json
{
  "source": "./source",
  "finish": "./finish",
  "token": "67dd867d******************************2ebcf536f9bc49e",
  "url": "https://www.noti***************9d91494e9a04fd9964ade210",
  "floor": ["asdafa","asd","kmos 2"],
  "primeKey": "订单编号",
  "info": [
    {
      "Execl": "编号",
      "Notion": "订单编号",
      "Type": "str"
    },
    {
      "Execl": "类型",
      "Notion": "订单类型",
      "Type": "select"
    },
    {
      "Execl": "分类",
      "Notion": "课程分类",
      "Type": "select"
    },
    {
      "Execl": "成单人",
      "Notion": "成单人",
      "Type": "str"
    },
    {
      "Execl": "头像",
      "Notion": "微信头像",
      "Type": "str"
    },
    {
      "Execl": "昵称",
      "Notion": "微信昵称",
      "Type": "str"
    },
    {
      "Execl": "手机号",
      "Notion": "手机号",
      "Type": "str"
    },
    {
      "Execl": "学员",
      "Notion": "上课学员",
      "Type": "str"
    },
    {
      "Execl": "销售",
      "Notion": "所属销售",
      "Type": "str"
    },
    {
      "Execl": "课程id",
      "Notion": "课程id",
      "Type": "int"
    },
    {
      "Execl": "课程名称",
      "Notion": "课程名称",
      "Type": "str"
    },
    {
      "Execl": "班次id",
      "Notion": "班次id",
      "Type": "int"
    }
  ]
}
```
