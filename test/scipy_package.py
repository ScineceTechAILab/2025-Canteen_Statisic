import scipy
import scipy.special

def main():
    # 简单使用scipy的功能，验证基本功能是否正常
    print(f"Scipy版本: {scipy.__version__}")
    
    # 调用scipy.special中的一个函数进行测试
    test_value = scipy.special.jn(0, 1.0)  # 贝塞尔函数
    print(f"测试计算结果: {test_value}")
    print("测试完成，功能正常")

if __name__ == "__main__":
    main()
