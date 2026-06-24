"""
供应链智能表单生成 - 主入口

用法:
    python main.py validate <数据文件路径>                    # 核查数据明细表
    python main.py generate-inspection <数据文件路径>          # 生成请验单
    python main.py generate-inbound <数据文件路径>             # 生成入库单
    python main.py generate-outbound <数据文件路径>            # 生成出库单
    python main.py verify <数据文件路径> [--output 表单目录]    # 核验表单
    python main.py all <数据文件路径>                          # 执行全部流程
"""

import argparse
import sys
import os

# 获取脚本所在目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'src'))
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'src', 'tools'))


def resolve_path(file_path):
    """解析文件路径，支持相对路径和assets目录"""
    if os.path.isabs(file_path):
        return file_path
    
    # 先检查当前目录
    if os.path.exists(file_path):
        return os.path.abspath(file_path)
    
    # 再检查assets目录
    assets_path = os.path.join(SCRIPT_DIR, 'assets', file_path)
    if os.path.exists(assets_path):
        return assets_path
    
    return os.path.abspath(file_path)


def cmd_validate(args):
    """核查数据明细表"""
    from data_validator import DataValidator
    
    data_file = resolve_path(args.data_file)
    print(f"📊 开始核查数据明细表: {data_file}\n")
    
    validator = DataValidator(data_file)
    result = validator.validate_all()
    
    # 格式化输出
    print(f"\n{'='*50}")
    print(f"📋 核查报告")
    print(f"{'='*50}")
    print(f"总行数: {result['total_rows']}")
    print(f"有效行数: {result['valid_rows']}")
    print(f"错误数: {result['error_count']}")
    print(f"警告数: {result['warning_count']}")
    
    if result['errors']:
        print(f"\n❌ 错误详情:")
        for i, err in enumerate(result['errors'], 1):
            print(f"  {i}. 行{err['row']+1} [{err['rule']}]: {err['message']}")
    
    if result['warnings']:
        print(f"\n⚠️  警告详情:")
        for i, warn in enumerate(result['warnings'], 1):
            print(f"  {i}. 行{warn['row']+1} [{warn['rule']}]: {warn['message']}")
    
    if result['error_count'] == 0 and result['warning_count'] == 0:
        print(f"\n✅ 数据核查通过，未发现错误")
    
    return result


def cmd_generate_inspection(args):
    """生成请验单"""
    from inspection_form_generator import InspectionFormGenerator
    
    data_file = resolve_path(args.data_file)
    output_dir = resolve_path(args.output) if args.output else os.path.join(SCRIPT_DIR, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    
    template_file = os.path.join(SCRIPT_DIR, 'assets', '请验单模板.xlsx')
    
    print(f"📝 开始生成请验单")
    print(f"   数据文件: {data_file}")
    print(f"   输出目录: {output_dir}\n")
    
    generator = InspectionFormGenerator(data_file, output_dir, template_file)
    result = generator.generate_all()
    
    print(f"\n{'='*50}")
    print(f"📋 生成结果")
    print(f"{'='*50}")
    print(f"总物料数: {result['total']}")
    print(f"成功: {result['success']}")
    print(f"失败: {result['failed']}")
    
    if result['files']:
        print(f"\n✅ 生成的文件:")
        for f in result['files']:
            print(f"   - {f['material_code']} ({f['material_name']}): {os.path.basename(f['file_path'])}")
    
    return result


def cmd_generate_inbound(args):
    """生成入库单"""
    from inbound_form_generator import InboundFormGenerator
    
    data_file = resolve_path(args.data_file)
    output_dir = resolve_path(args.output) if args.output else os.path.join(SCRIPT_DIR, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    
    template_file = os.path.join(SCRIPT_DIR, 'assets', '入库单单模板.xlsx')
    
    print(f"📝 开始生成入库单")
    print(f"   数据文件: {data_file}")
    print(f"   输出目录: {output_dir}\n")
    
    generator = InboundFormGenerator(data_file, output_dir, template_file)
    result = generator.generate_all()
    
    print(f"\n{'='*50}")
    print(f"📋 生成结果")
    print(f"{'='*50}")
    print(f"生成入库单数: {result['total']}")
    print(f"涉及物料数: {result['materials_count']}")
    print(f"成功: {result['success']}")
    print(f"失败: {result['failed']}")
    
    return result


def cmd_generate_outbound(args):
    """生成出库单"""
    from outbound_form_generator import OutboundFormGenerator
    
    data_file = resolve_path(args.data_file)
    output_dir = resolve_path(args.output) if args.output else os.path.join(SCRIPT_DIR, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    
    template_file = os.path.join(SCRIPT_DIR, 'assets', '出库单单模板.xlsx')
    
    print(f"📝 开始生成出库单")
    print(f"   数据文件: {data_file}")
    print(f"   输出目录: {output_dir}\n")
    
    generator = OutboundFormGenerator(data_file, output_dir, template_file)
    result = generator.generate_all()
    
    print(f"\n{'='*50}")
    print(f"📋 生成结果")
    print(f"{'='*50}")
    print(f"生成出库单数: {result['total']}")
    print(f"涉及出库数: {result['outbound_count']}")
    print(f"成功: {result['success']}")
    print(f"失败: {result['failed']}")
    
    return result


def cmd_verify(args):
    """核验表单"""
    from form_validator import FormValidator
    
    data_file = resolve_path(args.data_file)
    output_dir = resolve_path(args.output) if args.output else os.path.join(SCRIPT_DIR, 'outputs')
    
    print(f"🔍 开始核验表单")
    print(f"   数据文件: {data_file}")
    print(f"   表单目录: {output_dir}\n")
    
    validator = FormValidator(data_file, output_dir)
    result = validator.validate_all()
    
    print(f"\n{'='*50}")
    print(f"📋 核验报告")
    print(f"{'='*50}")
    print(f"总检查项: {result['summary']['total_checks']}")
    print(f"通过: {result['summary']['total_passed']}")
    print(f"失败: {result['summary']['total_failed']}")
    print(f"警告: {result['summary']['total_warnings']}")
    
    return result


def cmd_all(args):
    """执行全部流程"""
    print("="*50)
    print("🚀 执行完整供应链表单生成流程")
    print("="*50 + "\n")
    
    # 1. 数据核查
    print("【步骤 1/4】数据核查")
    validate_result = cmd_validate(args)
    
    if validate_result['error_count'] > 0:
        print(f"\n⚠️  数据存在 {validate_result['error_count']} 个错误，建议修正后再生成表单")
        answer = input("是否继续生成？(y/n): ")
        if answer.lower() != 'y':
            return
    
    # 2. 生成请验单
    print("\n【步骤 2/4】生成请验单")
    cmd_generate_inspection(args)
    
    # 3. 生成入库单
    print("\n【步骤 3/4】生成入库单")
    cmd_generate_inbound(args)
    
    # 4. 生成出库单
    print("\n【步骤 4/4】生成出库单")
    cmd_generate_outbound(args)
    
    print("\n" + "="*50)
    print("✅ 全部流程执行完成！")
    print("="*50)


def main():
    parser = argparse.ArgumentParser(description='供应链智能表单生成工具')
    subparsers = parser.add_subparsers(dest='command', help='可用命令')
    
    # validate 命令
    p_validate = subparsers.add_parser('validate', help='核查数据明细表')
    p_validate.add_argument('data_file', help='数据明细表文件路径')
    p_validate.set_defaults(func=cmd_validate)
    
    # generate-inspection 命令
    p_insp = subparsers.add_parser('generate-inspection', help='生成请验单')
    p_insp.add_argument('data_file', help='数据明细表文件路径')
    p_insp.add_argument('--output', '-o', help='输出目录')
    p_insp.set_defaults(func=cmd_generate_inspection)
    
    # generate-inbound 命令
    p_inbound = subparsers.add_parser('generate-inbound', help='生成入库单')
    p_inbound.add_argument('data_file', help='数据明细表文件路径')
    p_inbound.add_argument('--output', '-o', help='输出目录')
    p_inbound.set_defaults(func=cmd_generate_inbound)
    
    # generate-outbound 命令
    p_outbound = subparsers.add_parser('generate-outbound', help='生成出库单')
    p_outbound.add_argument('data_file', help='数据明细表文件路径')
    p_outbound.add_argument('--output', '-o', help='输出目录')
    p_outbound.set_defaults(func=cmd_generate_outbound)
    
    # verify 命令
    p_verify = subparsers.add_parser('verify', help='核验表单与数据一致性')
    p_verify.add_argument('data_file', help='数据明细表文件路径')
    p_verify.add_argument('--output', '-o', help='表单输出目录')
    p_verify.set_defaults(func=cmd_verify)
    
    # all 命令
    p_all = subparsers.add_parser('all', help='执行全部流程（核查+生成）')
    p_all.add_argument('data_file', help='数据明细表文件路径')
    p_all.add_argument('--output', '-o', help='输出目录')
    p_all.set_defaults(func=cmd_all)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(1)
    
    args.func(args)


if __name__ == '__main__':
    main()
