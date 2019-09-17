import autoit
import time


class MouseControl(object):
    """鼠标相关操作"""

    def __init__(self):
        pass

    @classmethod
    def click1(cls,title, text, x, y, button="", clicks=1):
        """
        :description 执行鼠标点击操作
        :return：
        """
        # 使用相对坐标的情况
        pos = autoit.win_get_pos(title, text=text)
        autoit.mouse_click(button, x + pos[0], y + pos[1], clicks=clicks)

    @classmethod
    def click(cls,lis,button="left",clicks=1):
        # 使用绝对坐标的情况
        autoit.mouse_click(button,lis[0],lis[1],clicks)

    @classmethod
    def move(cls,lis):
        """
        :description 移动鼠标指针
        :return:
        """
        # pos = autoit.win_get_pos(title, text=text)
        # autoit.mouse_click(button, x + pos[0], y + pos[1], clicks=clicks)
        # 使用绝对位置的情况
        autoit.mouse_move(lis[0], lis[1])

    @classmethod
    def drag(cls,lis):
        """
        :description 执行鼠标点击并拖动操作
        :return:
        """
        # pos = autoit.win_get_pos(title, text=text)
        # autoit.mouse_click_drag(x1 + pos[0], y1 + pos[1], x2 + pos[0], y2 + pos[1])
        # 使用绝对位置的情况
        autoit.mouse_click_drag(lis[0], lis[1], lis[2], lis[3])

    @classmethod
    def up(cls):
        """
        :description 执行鼠标当前位置的释放事件
        :return:
        """
        autoit.mouse_up()


    @classmethod
    def down(cls):
        """
        :description 执行鼠标当前位置的按下事件
        :return:
        """
        autoit.mouse_down()


    @classmethod
    def wheel(cls,direction="up"):
        """
        :description 执行鼠标滚轮向上或向下滚动事件
        :return:
        """
        autoit.mouse_wheel(direction)

    @classmethod
    def getCursor(cls):
        """
        :description
        :return:  返回当前鼠标光标的ID
        """
        ret = autoit.mouse_get_cursor()
        return ret

    @classmethod
    def getCurrentPos(cls):
        """
        :description
        :return:  返回当前鼠标所在的x,y坐标
        """
        ret = autoit.mouse_get_pos()
        return ret


if __name__ == "__main__":
    pass
