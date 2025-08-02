from jinja2 import Environment, BaseLoader

def main():
    # 定义一个简单的 Jinja2 模板
    template = Environment(loader=BaseLoader()).from_string("Hello, {{ name }}!")
    
    # 渲染模板
    result = template.render(name="Jinja2 打包测试")
    
    # 输出结果（验证功能正常）
    print(result)
    
    # 写入文件（验证打包后能否正常操作文件）
    with open("jinja_output.txt", "w", encoding="utf-8") as f:
        f.write(result)
    print("结果已写入 jinja_output.txt")

if __name__ == "__main__":
    main()