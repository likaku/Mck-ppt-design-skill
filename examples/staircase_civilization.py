#!/usr/bin/env python3
"""Example: Three Stages of Human Civilization — Staircase Evolution (#15).

Demonstrates:
  - PNG icon support (icon_factory_gear, icon_circuit_chip, icon_ai_brain)
  - Single-line detail_rows (no bullet prefix)
  - Narrative-style complete sentences

Output: examples/civilization_staircase.pptx
"""
import os
import sys

# Add parent dir to path so we can import mck_ppt
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mck_ppt import MckEngine
from mck_ppt.core import full_cleanup

ICON_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        'assets', 'icons')


def main():
    # Ensure icons exist — generate if missing
    required_icons = ['icon_factory_gear.png', 'icon_circuit_chip.png', 'icon_ai_brain.png']
    missing = [ic for ic in required_icons if not os.path.isfile(os.path.join(ICON_DIR, ic))]
    if missing:
        print(f'Generating missing icons: {missing}')
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                        'scripts'))
        import generate_icons
        for name, fn in generate_icons.ALL_ICONS:
            if name in missing:
                path = os.path.join(ICON_DIR, name)
                fn().save(path, 'PNG')
                print(f'  Generated: {path}')

    eng = MckEngine(total_slides=1)

    eng.pyramid(
        title='人类文明演进：从蒸汽机到超级智能的三次跃迁',
        levels=[
            (
                '工业时代 1760–1970',
                '"蒸汽与电力驱动的规模化生产"\n机械动力替代人力，人类第一次突破体力极限',
                os.path.join(ICON_DIR, 'icon_factory_gear.png'),
            ),
            (
                '信息时代 1970–2020',
                '"数据与互联网连接的全球化协作"\n计算与通信替代重复脑力，人类第一次突破时空限制',
                os.path.join(ICON_DIR, 'icon_circuit_chip.png'),
            ),
            (
                '智能时代 2020–未来',
                '"AI与自主系统驱动的认知革命"\n通用智能替代决策判断，人类第一次突破智力极限',
                os.path.join(ICON_DIR, 'icon_ai_brain.png'),
            ),
        ],
        detail_rows=[
            ('核心技术', [
                ['蒸汽机、电力、内燃机、流水线——用机器取代手工，实现标准化大规模生产'],
                ['晶体管、互联网、移动通信、云计算——用比特取代原子传递信息，实现全球实时协作'],
                ['大语言模型、具身智能、自主Agent——用神经网络模拟推理，实现机器自主认知与决策'],
            ]),
            ('社会变革', [
                ['城市化、中产阶级崛起、民族国家体系形成——人类从农村迁入工厂，劳动生产率提升50倍'],
                ['全球化、平台经济、知识工作者涌现——信息不对称被消除，个体可以连接全世界'],
                ['人机协作、职业重构、智能体经济——AI成为新的生产要素，创造力成为核心竞争力'],
            ]),
        ],
        source='Source: Schwab, K. "The Fourth Industrial Revolution" (2016); OpenAI "Planning for AGI" (2023)',
    )

    outdir = os.path.dirname(os.path.abspath(__file__))
    outpath = os.path.join(outdir, 'civilization_staircase.pptx')
    eng.save(outpath)
    full_cleanup(outpath)
    print(f'\n✅ PPT saved: {outpath}')


if __name__ == '__main__':
    main()
