# coding: UTF-8

from PIL import Image
import numpy as np
import openpyxl
import matplotlib
font = {"family":"AppleGothic"}
matplotlib.rc('font', **font)
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.path import Path
from matplotlib.spines import Spine
from matplotlib.projections.polar import PolarAxes
from matplotlib.projections import register_projection

# レーダーチャートを描く
def radar_factory(num_vars, frame='polygon'):
    # 質問項目を上部から時計回りに配置する
    theta = np.linspace(0, -2*np.pi, num_vars, endpoint=False)
    theta += np.pi/2

    def draw_poly_patch(self):
        verts = unit_poly_verts(theta)
        return plt.Polygon(verts, closed=True, edgecolor='k')

    def draw_circle_patch(self):
        # unit circle centered on (0.5, 0.5)
        return plt.Circle((0.5, 0.5), 0.5)

    patch_dict = {'polygon': draw_poly_patch, 'circle': draw_circle_patch}
    if frame not in patch_dict:
        raise ValueError('unknown value for `frame`: %s' % frame)

    class RadarAxes(PolarAxes):
        name = 'radar'
        # 指定した点を結ぶ線分を1つ使用する
        RESOLUTION = 1
        draw_patch = patch_dict[frame]
        # チャートを塗りつぶす
        def fill(self, *args, **kwargs):
            closed = kwargs.pop('closed', True)
            return super(RadarAxes, self).fill(closed=closed, *args, **kwargs)
        # チャートの枠線を描く
        def plot(self, *args, **kwargs):
            lines = super(RadarAxes, self).plot(*args, **kwargs)
            for line in lines:
                self._close_line(line)

        def _close_line(self, line):
            x, y = line.get_data()
            # FIXME: markers at x[0], y[0] get doubled-up
            if x[0] != x[-1]:
                x = np.concatenate((x, [x[0]]))
                y = np.concatenate((y, [y[0]]))
                line.set_data(x, y)

        def set_varlabels(self, labels):
            self.set_thetagrids(np.degrees(theta), labels)

        def _gen_axes_patch(self):
            return self.draw_patch()

        def _gen_axes_spines(self):
            if frame == 'circle':
                return PolarAxes._gen_axes_spines(self)
            # The following is a hack to get the spines (i.e. the axes frame)
            # to draw correctly for a polygon frame.

            # spine_type must be 'left', 'right', 'top', 'bottom', or `circle`.
            spine_type = 'circle'
            verts = unit_poly_verts(theta)
            # close off polygon by repeating first vertex
            verts.append(verts[0])
            path = Path(verts)

            spine = Spine(self, spine_type, path)
            spine.set_transform(self.transAxes)
            return {'polar': spine}

    register_projection(RadarAxes)
    return theta

# 多角形の頂点を返す((0.5, 0.5)中心の単位円が基準）
def unit_poly_verts(theta):
    x0, y0, r = [0.5] * 3
    verts = [(r*np.cos(t) + x0, r*np.sin(t) + y0) for t in theta]
    return verts

# ここはExcelデータに置き換える
def import_data():
    data = [
        ['01', '02', '03', '04', '05', '06', '07', '08'],
        ('Graph', [
            [7.00, 3.83, 5.00, 5.50, 3.83, 4.00, 4.33, 4.00],
            [4.00, 4.50, 3.50, 5.00, 4.17, 3.20, 4.33, 5.00],
            [0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00]])
    ]
    return data


if __name__ == '__main__':
    N = 8
    theta = radar_factory(N, frame='polygon')

    data = import_data()
    spoke_labels = data.pop(0)

    figure, axes = plt.subplots(figsize=(8, 8), nrows=1, ncols=1,
                             subplot_kw=dict(projection='radar'))

    # figure.subplots_adjust(wspace=0.25, hspace=0.20, top=0.85, bottom=0.05)

    colors = ['b', 'r', 'g', 'm', 'y']
    # データからチャートを生成する
    for ax, (title, case_data) in zip([axes], data):
        # 目盛
        ax.set_rgrids([1.0, 3.0, 5.0, 7.0])
        # 塗りつぶしなしの線だけでチャート描画
        for d, color in zip(case_data, colors):
            ax.plot(theta, d, color=color)
        ax.set_varlabels(spoke_labels)
    # 凡例を配置
    labels = ('Grade1-1', 'Grade1-2', 'Grade2-1', 'Grade2-2', 'Grade3-1')
    legend = axes.legend(labels, loc=(0.9, .95),
                       labelspacing=0.1, fontsize='small')
    # 作成したチャートを画像出力
    plt.savefig('graph.png')
    # 出力したgraph.pngをout.xlsxに貼り付ける
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]
    img = openpyxl.drawing.image.Image('graph.png')
    ws.add_image(img, 'B12')
    wb.save('out.xlsx')
    # pil_img = Image.fromarray(plt.savefig('graph.png'))
    # pil_img.save('data/dst/lena_square_save.png')
