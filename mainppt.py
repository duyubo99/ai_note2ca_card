from core.ppt_generator import generate_ppt

def main():
    """主函数，用于生成PPT"""
    # 生成PPT
    generated_files = generate_ppt()
    
    # 打印结果
    print(f"共生成 {len(generated_files)} 个PPT文件:")
    for file_path in generated_files:
        print(f"  - {file_path}")
    
    return generated_files

if __name__ == "__main__":
    main()
