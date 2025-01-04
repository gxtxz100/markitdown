import os
from pathlib import Path
from markitdown import MarkItDown
from tqdm import tqdm

def convert_to_markdown(input_dir: str) -> None:
    """
    批量将文件夹中的文档转换为Markdown文件
    
    Args:
        input_dir: 输入文件夹路径
    """
    # 直接使用默认设置创建实例
    converter = MarkItDown()
    
    # 支持的文件扩展名
    supported_extensions = {
        '.docx', '.pdf', '.pptx', '.xlsx', '.xls',
        '.html', '.htm', '.msg', '.jpg', '.jpeg',
        '.png', '.json', '.xml', '.csv', '.zip'
    }
    
    # 首先统计需要转换的文件总数
    total_files = sum(1 for root, _, files in os.walk(input_dir)
                     for file in files
                     if os.path.splitext(file)[1].lower() in supported_extensions)
    
    # 使用tqdm创建进度条
    with tqdm(total=total_files, desc="转换进度") as pbar:
        for root, _, files in os.walk(input_dir):
            for file in files:
                input_path = os.path.join(root, file)
                file_extension = os.path.splitext(file)[1].lower()
                
                if file_extension in supported_extensions:
                    try:
                        print(f"正在转换: {input_path}")
                        
                        # 转换文件
                        result = converter.convert(input_path)
                        
                        # 检查转换结果
                        if not result.text_content.strip():
                            print(f"警告: {input_path} 未能提取到文字内容")
                            pbar.update(1)
                            continue
                            
                        # 写入Markdown文件
                        output_filename = os.path.splitext(file)[0] + '.md'
                        output_path = os.path.join(root, output_filename)
                        with open(output_path, 'w', encoding='utf-8') as f:
                            f.write(result.text_content)
                            
                        print(f"转换成功: {output_path}")
                    except Exception as e:
                        print(f"转换失败 {input_path}: {str(e)}")
                    finally:
                        pbar.update(1)
                else:
                    print(f"不支持的文件类型: {input_path}")
                    pbar.update(1)

def main():
    print("欢迎使用文档转Markdown批量转换工具")
    
    input_dir = input("请输入源文件夹路径: ").strip()
    
    if not os.path.isdir(input_dir):
        print("输入的不是有效的文件夹路径，程序退出")
        return
    
    print(f"\n开始转换文件夹: {input_dir}\n")
    convert_to_markdown(input_dir)
    print("\n转换完成!")

if __name__ == "__main__":
    main() 