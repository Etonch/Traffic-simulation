import numpy as np
import math
import random
import xlrd
import xlsxwriter
import os

class Cars():  # 车类，包括每个车辆数量，速度，位置,需要其他参数可以加
    def __init__(self, car_num, car_length):
        self.car_num = car_num
        self.car_speed = np.zeros(car_num)  # 车速，一维数组，初始置0
        self.car_location = []  # 车的坐标，每个坐标用x和y两个值表示(坐标以车头为锚点）
        self.car_type = np.zeros(car_num)  # 车辆种类，人工驾驶车为1，自动驾驶车为2，初始值置0
        self.car_length = car_length
        self.car_acc = np.zeros(car_num)
        self.car_lc = []  # 车辆换道标号，记录车辆每次即将变向的车道号，初始置-1
        for _ in range(car_num):
            self.car_lc.append(-1)


class Road():
    def __init__(self, cells_shape):
        self.lane_length = cells_shape[0]  # 车道长度
        self.lane_num = cells_shape[1]  # 车道数量
        # 车道种类，普通车道为0，人工驾驶专用道为1，自动驾驶车专用道为2
        self.lane_type = np.zeros(self.lane_num)
        self.flux = np.zeros(self.lane_num)


# 元胞自动机类，继承车类和道路类
class Automata(Cars, Road):
    def __init__(self, car_num, car_length, cells_shape, cav_proportion, elusive_lane, j):
        Cars.__init__(self, car_num, car_length)
        Road.__init__(self, cells_shape)
        self.cells = np.zeros(cells_shape)
        for i in range(self.lane_num):
            self.lane_type[i] = int(elusive_lane[i])  # 专用道标号

        self.timer = 0  # 计时器，记录迭代次数
        self.cav_proportion = cav_proportion  # cav占有率
        self.MV_num = car_num-int(car_num*cav_proportion)
        self.CAV_num = int(car_num*cav_proportion)
        self.lane_avspeed = np.zeros((4, 3600))
        self.MV_avspeed = []
        self.CAV_avspeed = []
        self.head = np.zeros(4)
        self.rear = np.zeros(4)
        self.num = j
        self.elusive_lane=elusive_lane
        self.change_in_nums=np.zeros(4)
        self.change_out_nums=np.zeros(4)
        self.veh_num = np.zeros((4, 3600))

    # 车辆初始化，考虑车身长度，将道路分为长度为车辆长度的小段，作为车辆的可能生成点，并从中选择，运行过程中舍去前n次迭代过程即可认为随机

    def car_distribute(self):
        ret = []
        if self.cav_proportion < 0.5:
            list = []
            for x in range(int(self.lane_length/self.car_length)):
                for y in range(self.lane_num):
                    # 生成位置限制，只能在普通车道和CAV专用道生成
                    if self.lane_type[y] == 2 or self.lane_type[y] == 0:
                        list.append([x*self.car_length-1, y])
            self.car_location = random.sample(list, self.CAV_num)

            list = []
            for x in range(int(self.lane_length/self.car_length-1)):
                for y in range(self.lane_num):
                    if self.lane_type[y] == 1 or self.lane_type[y] == 0:
                        list.append([x*self.car_length-1, y])

            for i in list:
                if i not in self.car_location:
                    ret.append(i)

            tmp = random.sample(ret, self.MV_num)
            for i in range(self.MV_num):
                self.car_location.append(tmp[i])

            for i in range(self.car_num):
                self.cells[self.car_location[i][0]
                           ][self.car_location[i][1]] = 1
            list1 = self.car_location
            list1 = sorted(list1, key=(lambda x: [x[1], x[0]]))
            for i in range(self.CAV_num):
                m = list1.index(self.car_location[i])
                self.car_type[m] = 2
            for i in range(self.CAV_num, self.car_num):
                m = list1.index(self.car_location[i])
                self.car_type[m] = 1
            self.car_location = list1

        else:
            list = []
            for x in range(int(self.lane_length / self.car_length)):
                for y in range(self.lane_num):
                    # 生成位置限制，只能在普通车道和CAV专用道生成
                    if self.lane_type[y] == 1 or self.lane_type[y] == 0:
                        list.append([x * self.car_length - 1, y])
            self.car_location = random.sample(list, self.MV_num)

            list = []
            for x in range(int(self.lane_length / self.car_length - 1)):
                for y in range(self.lane_num):
                    if self.lane_type[y] == 2 or self.lane_type[y] == 0:
                        list.append([x * self.car_length - 1, y])

            for i in list:
                if i not in self.car_location:
                    ret.append(i)

            tmp = random.sample(ret, self.CAV_num)
            for i in range(self.CAV_num):
                self.car_location.append(tmp[i])

            for i in range(self.car_num):
                self.cells[self.car_location[i][0]
                           ][self.car_location[i][1]] = 1
            list1 = self.car_location
            list1 = sorted(list1, key=(lambda x: [x[1], x[0]]))
            for i in range(self.MV_num):
                m = list1.index(self.car_location[i])
                self.car_type[m] = 1
            for i in range(self.MV_num, self.car_num):
                m = list1.index(self.car_location[i])
                self.car_type[m] = 2
            self.car_location = list1

        # print(self.car_location)
        # print(self.car_type)

        rear = self.car_num
        for i in reversed(range(self.lane_num)):
            flag = 1
            head = rear - 1
            self.head[i] = head
            for m in reversed(range(head+1)):  # 不考虑车道只有一辆车的情况
                if self.car_location[m][1] == i - 1:
                    rear = m + 1
                    flag = 0
                    break
            if flag:
                rear = 0
            self.rear[i] = rear

    def update_speed(self):
        bmax = 3                   # 给定参数
        t = 0.1
        vmax = 66
        amv = 1
        amax = 3
        T = 1.8
        dmax = 600                  # 最大通信距离
        kp = 0.45
        kd = 0.25
        k1 = 0.23
        k2 = 0.07
        tc = 0.6
        ta = 1.1
        rear = self.car_num
        cav_sum = 0
        mv_sum = 0
        sumspeed = 0
        for i in reversed(range(self.lane_num)):
            flag = 1
            ff = 0
            head = rear - 1
            self.head[i] = head
            for m in reversed(range(head)):  # 不考虑车道只有一辆车的情况
                if self.car_location[m][1] == i - 1:
                    rear = m + 1
                    flag = 0
                    break
            if flag:
                rear = 0
            self.rear[i] = rear
            # print(self.head,self.rear)

            for j in reversed(range(rear, head+1)):  # j为当前车序号
                # print(i,j,ff,rear,head)
                if j == head:  # 队头的车
                    if self.car_type[j] == 1:  # MV
                        d = self.car_location[rear][0]+self.lane_length - \
                            self.car_location[j][0]-self.car_length
                        dl = self.car_location[rear+1][0] - \
                            self.car_location[rear][0]-self.car_length
                        vl = self.car_speed[rear]
                        al = self.car_acc[rear]

                        vsafe = -bmax + \
                            math.sqrt(bmax * bmax + (vl-al) * (vl-al) + 2 * bmax * d)
                        vanti = min(dl, vl-al+1, vmax)
                        if self.car_type[rear] == 1:
                            g = 1.8
                        else:
                            g = 2.4
                        danti = min(
                            (d + vanti + self.car_length) / (1 + g), d + vanti)
                        v = int(
                            min(self.car_speed[j] + amv, vmax, danti, vsafe))
                        if v <= danti / T:  # MV随机慢化
                            p = np.array([0.1, 0.9])
                        else:
                            p = np.array([0.1 + 0.85 / (1 + math.e ** (10 * (30 - v))),
                                          0.9 - 0.85 / (1 + math.e ** (10 * (30 - v)))])
                        v = np.random.choice([max(v - 1, 0), v], p=p.ravel())

                    if self.car_type[j] == 2:
                        d = self.car_location[rear][0] + self.lane_length - \
                            self.car_location[j][0]-self.car_length
                        dl = self.car_location[rear + 1][0] - \
                            self.car_location[rear][0]-self.car_length
                        vl = self.car_speed[rear]
                        al = self.car_acc[rear]

                        vsafe = -bmax*t + \
                            math.sqrt(bmax * bmax * t*t +
                                      (vl-al)*(vl-al) + 2 * bmax * d)
                        if self.car_type[rear] == 1:
                            g = 0.9
                            acav = int(max(min(
                                amax, k1 * (d - ta * self.car_speed[j]) + k2 * (vl - self.car_speed[j])), -bmax))
                        else:
                            g = 0
                            acav = int(max(-bmax, min(amax, kp * (d - tc * self.car_speed[j]) + kd * (
                                vl - self.car_speed[j] - tc * self.car_acc[j]))))
                        
                        if self.car_type[rear]==2:
                            sum = 0
                            n = 0
                            if self.car_location[j][0] + dmax  < self.lane_length:
                                vanti = min(dl, vl-al+1, vmax)
                            else:
                                for m in range(rear, head):
                                    if self.car_type[m] == 2 and \
                                            (self.car_location[m][0] <= self.car_location[j][0] + dmax - self.lane_length or
                                             self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] + dmax):
                                        sum += self.car_speed[m]
                                        n = n + 1
                            
                                if n!=0:
                                    vanti = min(dl, vl, vmax, sum / n)
                                else:
                                    vanti = min(dl, vl-al+1, vmax)
                        else:
                            vanti = min(dl, vl-al+1, vmax)

                        danti = min(
                            (d + vanti + self.car_length) / (1 + g), d + vanti)
                        v = int(
                            max(min(self.car_speed[j] + acav, vmax, danti, vsafe, d+vl),0))

                    self.car_acc[j] = v - self.car_speed[j]
                    self.car_speed[j] = v
                    if self.timer >= 2000:
                        if (self.car_type[j] == 2):
                            cav_sum += v
                        elif (self.car_type[j] == 1):
                            mv_sum += v
                    sumspeed += v
                    continue

                if j != head:  #非队头的车
                    if self.car_type[j] == 1:  # MV
                        d = self.car_location[j+1][0] - \
                            self.car_location[j][0]-self.car_length
                        if j+2 <= head:
                            dl = self.car_location[j+2][0] - \
                                self.car_location[j+1][0]-self.car_length
                        else:
                            dl = self.car_location[rear][0]+self.lane_length - \
                                self.car_location[j+1][0]-self.car_length
                        vl = self.car_speed[j+1]
                        al = self.car_acc[j+1]
                        vsafe = -bmax + \
                            math.sqrt(bmax * bmax + (vl-al) * (vl-al) + 2 * bmax * d)
                        vanti = min(dl, vl-al+1, vmax)
                        if self.car_type[j+1] == 1:
                            g = 1.8
                        else:
                            g = 2.4
                        danti = min(
                            (d + vanti + self.car_length) / (1 + g), d + vanti)
                        v = int(
                            min(self.car_speed[j] + amv, vmax, danti, vsafe))
                        if v <= danti / T:  # MV随机慢化
                            p = np.array([0.1, 0.9])
                        else:
                            p = np.array([0.1 + 0.85 / (1 + math.e ** (10 * (30 - v))),
                                          0.9 - 0.85 / (1 + math.e ** (10 * (30 - v)))])
                        v = np.random.choice([max(v - 1, 0), v], p=p.ravel())

                    if self.car_type[j] == 2:
                        d = self.car_location[j + 1][0] - \
                            self.car_location[j][0]-self.car_length
                        if j + 2 <= head:
                            dl = self.car_location[j + 2][0] - \
                                self.car_location[j + 1][0]-self.car_length
                        else:
                            dl = self.car_location[rear][0] + self.lane_length - \
                                self.car_location[j + 1][0]-self.car_length
                        vl = self.car_speed[j + 1]
                        al = self.car_acc[j + 1]

                        vsafe = -bmax*t + \
                            math.sqrt(bmax * bmax * t*t +
                                      (vl-al)*(vl-al) + 2 * bmax * d)

                        if self.car_type[j+1] == 1:
                            g = 0.9
                            acav = int(max(min(
                                amax, k1 * (d - ta * self.car_speed[j]) + k2 * (vl - self.car_speed[j])), -bmax))
                        else:
                            g = 0
                            acav = int(max(-bmax, min(amax, kp * (d - tc * self.car_speed[j]) + kd * (
                                vl - self.car_speed[j] - tc * self.car_acc[j]))))

                        if self.car_type[j+1]==2:
                            sum = 0
                            n = 0
                            if self.car_location[j][0] + dmax < self.lane_length:
                                for m in range(j+1, head):
                                    if self.car_type[m] == 2 and \
                                            (self.car_location[m][0] <= self.car_location[j][
                                                0] + dmax - self.lane_length or
                                            self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] +dmax):
                                        sum += self.car_speed[m]
                                        n = n + 1
                            else:
                                for m in range(rear, head):
                                    if self.car_type[m] == 2 and \
                                            (self.car_location[m][0] <= self.car_location[j][
                                                0] + dmax - self.lane_length or
                                             self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] +dmax):
                                        sum += self.car_speed[m]
                                        n = n + 1
                                if n!=0:
                                    vanti = min(dl, vl, vmax, sum / n)
                                else:
                                    vanti = min(dl, vl-al+1, vmax)
                        else:
                            vanti = min(dl, vl-al+1, vmax)
                        danti = min(
                            (d + vanti + self.car_length) / (1 + g), d + vanti)
                        v = int(
                            max(min(self.car_speed[j] + acav, vmax, danti, vsafe, d+vl),0))
                    self.car_acc[j] = v - self.car_speed[j]
                    self.car_speed[j] = v
                    if self.timer >= 2000:
                        if (self.car_type[j] == 2):
                            cav_sum += v
                        elif (self.car_type[j] == 1):
                            mv_sum += v
                    sumspeed += v
                    continue

            j = int(self.head[i])
            if self.car_location[rear][0]+self.car_speed[rear]-self.car_location[j][0]-self.car_speed[j]+self.lane_length < self.car_length:
                if self.car_type[j] == 1:  # MV
                    d = self.car_location[rear][0] + self.lane_length - \
                        self.car_location[j][0] - self.car_length
                    dl = self.car_location[rear + 1][0] - \
                        self.car_location[rear][0] - self.car_length
                    vl = self.car_speed[rear]
                    al = self.car_acc[rear]
                    vsafe = -bmax + \
                        math.sqrt(bmax * bmax + (vl-al)*(vl-al) + 2 * bmax * d)
                    vanti = min(dl, vl, vmax)
                    if self.car_type[rear] == 1:
                        g = 1.8
                    else:
                        g = 2.4
                    danti = min((d + vanti + self.car_length) /
                                (1 + g), d + vanti)
                    v = int(min(self.car_speed[j] + amv, vmax, danti, vsafe))
                    if v <= danti / T:  # MV随机慢化
                        p = np.array([0.1, 0.9])
                    else:
                        p = np.array([0.1 + 0.85 / (1 + math.e ** (10 * (30 - v))),
                                      0.9 - 0.85 / (1 + math.e ** (10 * (30 - v)))])
                    v = np.random.choice([max(v - 1, 0), v], p=p.ravel())

                if self.car_type[j] == 2:
                    d = self.car_location[rear][0] + self.lane_length - \
                        self.car_location[j][0] - self.car_length
                    dl = self.car_location[rear + 1][0] - \
                        self.car_location[rear][0] - self.car_length
                    vl = self.car_speed[rear]
                    al = self.car_acc[rear]
                    vsafe = -bmax * t + \
                        math.sqrt(bmax * bmax * t * t + (vl-al) * (vl-al) + 2 * bmax * d)
                    if self.car_type[rear] == 1:
                        g = 0.9
                        acav = int(
                            max(min(amax, k1 * (d - ta * self.car_speed[j]) + k2 * (vl - self.car_speed[j])), -bmax))
                    else:
                        g = 0
                        acav = int(max(-bmax, min(amax, kp * (d - tc * self.car_speed[j]) + kd * (
                            vl - self.car_speed[j] - tc * self.car_acc[j]))))

                    if self.car_type[rear]==2:
                        sum = 0
                        n = 0
                        if self.car_location[j][0] + dmax < self.lane_length:
                            vanti = min(dl, vl-al+1, vmax)
                        else:
                            for m in range(rear, head):
                                if self.car_type[m] == 2 and \
                                        (self.car_location[m][0] <= self.car_location[j][0] + dmax - self.lane_length or
                                        self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] + dmax):
                                    sum += self.car_speed[m]
                                    n = n + 1
                            
                            if n!=0:
                                vanti = min(dl, vl, vmax, sum / n)
                            else:
                                vanti = min(dl, vl-al+1, vmax)
                    else:
                        vanti = min(dl, vl-al+1, vmax)
                    danti = min((d + vanti + self.car_length) /
                                (1 + g), d + vanti)
                    v = int(
                            max(min(self.car_speed[j] + acav, vmax, danti, vsafe, d+vl),0))

                if self.timer >= 2000: 
                    if (self.car_type[j] == 2):
                        cav_sum -= self.car_speed[j]
                        sumspeed -= self.car_speed[j]
                        self.car_acc[j] = v-(self.car_speed[j]-self.car_acc[j])
                        self.car_speed[j] = v
                        sumspeed += self.car_speed[j]
                        cav_sum += self.car_speed[j]
                    elif (self.car_type[j] == 1):
                        mv_sum -= self.car_speed[j]
                        sumspeed -= self.car_speed[j]
                        self.car_acc[j] = v - \
                            (self.car_speed[j] - self.car_acc[j])
                        self.car_speed[j] = v
                        sumspeed += self.car_speed[j]
                        mv_sum += self.car_speed[j]
                else:
                    self.car_acc[j] = v - (self.car_speed[j] - self.car_acc[j])
                    self.car_speed[j] = v

            j = int(self.head[i])-1
            while (self.car_location[j+1][0]+self.car_speed[j+1]-self.car_location[j][0]-self.car_speed[j] < self.car_length):
                if self.car_type[j] == 1:  # MV
                    d = self.car_location[j + 1][0] - \
                        self.car_location[j][0] - self.car_length
                    if j + 2 <= head:
                        dl = self.car_location[j + 2][0] - \
                            self.car_location[j + 1][0] - self.car_length
                    else:
                        dl = self.car_location[rear][0] + self.lane_length - self.car_location[j + 1][
                            0] - self.car_length
                    vl = self.car_speed[j + 1]
                    al = self.car_acc[j + 1]
                    vsafe = -bmax + \
                        math.sqrt(bmax * bmax + (vl-al) * (vl-al) + 2 * bmax * d)
                    vanti = min(dl, vl, vmax)
                    if self.car_type[j + 1] == 1:
                        g = 1.8
                    else:
                        g = 2.4
                    danti = min((d + vanti + self.car_length) /
                                (1 + g), d + vanti)
                    v = int(min(self.car_speed[j] + amv, vmax, danti, vsafe))
                    if v <= danti / T:  # MV随机慢化
                        p = np.array([0.1, 0.9])
                    else:
                        p = np.array([0.1 + 0.85 / (1 + math.e ** (10 * (30 - v))),
                                      0.9 - 0.85 / (1 + math.e ** (10 * (30 - v)))])
                    v = np.random.choice([max(v - 1, 0), v], p=p.ravel())

                if self.car_type[j] == 2:
                    d = self.car_location[j + 1][0] - \
                        self.car_location[j][0] - self.car_length
                    if j + 2 <= head:
                        dl = self.car_location[j + 2][0] - \
                            self.car_location[j + 1][0] - self.car_length
                    else:
                        dl = self.car_location[rear][0] + self.lane_length - self.car_location[j + 1][
                            0] - self.car_length
                    vl = self.car_speed[j + 1]
                    al = self.car_acc[j + 1]
                    vsafe = -bmax * t + \
                        math.sqrt(bmax * bmax * t * t + (vl-al)*(vl-al) + 2 * bmax * d)
                    if self.car_type[j + 1] == 1:
                        g = 0.9
                        acav = int(
                            max(min(amax, k1 * (d - ta * self.car_speed[j]) + k2 * (vl - self.car_speed[j])), -bmax))
                    else:
                        g = 0
                        acav = int(max(-bmax, min(amax, kp * (d - tc * self.car_speed[j]) + kd * (
                            vl - self.car_speed[j] - tc * self.car_acc[j]))))

                    if self.car_type[j+1]==2:
                        sum = 0
                        n = 0
                        if self.car_location[j][0] + dmax < self.lane_length:
                            for m in range(j+1, head):
                                if self.car_type[m] == 2 and \
                                        (self.car_location[m][0] <= self.car_location[j][
                                            0] + dmax - self.lane_length or
                                         self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] +dmax):
                                    sum += self.car_speed[m]
                                    n = n + 1
                        else:
                            for m in range(rear, head):
                                if self.car_type[m] == 2 and \
                                        (self.car_location[m][0] <= self.car_location[j][
                                            0] + dmax - self.lane_length or
                                         self.car_location[j][0]<self.car_location[m][0] <= self.car_location[j][0] +dmax):
                                    sum += self.car_speed[m]
                                    n = n + 1
                            
                            if n!=0:
                                vanti = min(dl, vl, vmax, sum / n)
                            else:
                                vanti = min(dl, vl-al+1, vmax)
                    else:
                        vanti = min(dl, vl-al+1, vmax)

                    danti = min((d + vanti + self.car_length) /
                                (1 + g), d + vanti)
                    v = int(
                            max(min(self.car_speed[j] + acav, vmax, danti, vsafe, d+vl),0))

                if self.timer >= 2000:
                    if (self.car_type[j] == 2):
                        cav_sum -= self.car_speed[j]
                        sumspeed -= self.car_speed[j]
                        self.car_acc[j] = v - \
                            (self.car_speed[j] - self.car_acc[j])
                        self.car_speed[j] = v
                        sumspeed += self.car_speed[j]
                        cav_sum += self.car_speed[j]
                    elif (self.car_type[j] == 1):
                        mv_sum -= self.car_speed[j]
                        sumspeed -= self.car_speed[j]
                        self.car_acc[j] = v - \
                            (self.car_speed[j] - self.car_acc[j])
                        self.car_speed[j] = v
                        sumspeed += self.car_speed[j]
                        mv_sum += self.car_speed[j]
                else:
                    self.car_acc[j] = v - (self.car_speed[j] - self.car_acc[j])
                    self.car_speed[j] = v
            if self.timer >= 2000:
                if head-rear+1 != 0:
                    self.lane_avspeed[i][self.timer -
                                         2000] = sumspeed / (head - rear + 1)
                else:
                    self.lane_avspeed[i][self.timer - 2000] = 0
            sumspeed = 0
        if self.timer >= 2000:
            if self.MV_num!=0:
                self.MV_avspeed.append(mv_sum/self.MV_num)
            else:
                self.MV_avspeed.append(0)
            if self.CAV_num!=0:
                self.CAV_avspeed.append(cav_sum / self.CAV_num)
            else:
                self.CAV_avspeed.append(0)

    def update_state(self):
        self.update_speed()
        for i in reversed(range(self.lane_num)):
            ff = 0
            j = int(self.head[i])
            while (self.rear[i]+ff <= j <= self.head[i]):
                if j == self.head[i]:
                    if self.car_location[j][0] + self.car_speed[j] < self.lane_length:
                        self.cells[int(self.car_location[j][0])
                                   ][int(self.car_location[j][1])] = 0
                        self.car_location[j][0] = self.car_location[j][0] + \
                            self.car_speed[j]
                        self.cells[int(self.car_location[j][0])
                                   ][int(self.car_location[j][1])] = 1

                    elif self.car_location[j][0] + self.car_speed[j] >= self.lane_length:

                        self.cells[int(self.car_location[j][0])
                                   ][int(self.car_location[j][1])] = 0
                        self.car_location[j][0] = self.car_location[j][0] + \
                            self.car_speed[j] - self.lane_length
                        self.cells[int(self.car_location[j][0])
                                   ][int(self.car_location[j][1])] = 1
                        x = self.car_location[j][0]
                        speed = self.car_speed[j]
                        acc = self.car_acc[j]
                        typ = self.car_type[j]
                        for m in range(j, int(self.rear[i]), -1):
                            self.car_location[m][0] = self.car_location[m - 1][0]
                            self.car_speed[m] = self.car_speed[m - 1]
                            self.car_acc[m] = self.car_acc[m - 1]
                            self.car_type[m] = self.car_type[m - 1]
                        self.car_location[int(self.rear[i])][0] = x
                        self.car_speed[int(self.rear[i])] = speed
                        self.car_acc[int(self.rear[i])] = acc
                        self.car_type[int(self.rear[i])] = typ
                        j += 1
                        ff += 1

                elif j != self.head[i]:
                    self.cells[int(self.car_location[j][0])
                               ][int(self.car_location[j][1])] = 0
                    self.car_location[j][0] = self.car_location[j][0] + \
                        self.car_speed[j]
                    self.cells[int(self.car_location[j][0])
                               ][int(self.car_location[j][1])] = 1

                j -= 1
            if self.timer>=2000:
                self.flux[i] += ff    #统计流量

        self.timer += 1

    def lane_change(self):
        def other_front(i, y):  # i为车辆序号，y为潜在可变车道号，计算其他车道前车距离
            flag = 1
            for j in range(int(self.car_location[i][0] + 1), self.lane_length):
                if self.cells[j][y] == 1:
                    d3 = j - self.car_location[i][0] - self.car_length
                    flag = 0
                    break
            if flag:
                for j in range(0, int(self.car_location[i][0]+1)):
                    if self.cells[j][y] == 1:
                        d3 = j - self.car_location[i][0] - \
                            self.car_length + self.lane_length
                        flag = 0
                        break
            if flag:
                d3 = self.lane_length
            return d3

        def other_back(i, y):  # i为车辆序号，y为潜在可变车道号，计算其他车道后车距离
            flag = 1
            for j in reversed(range(0, int(self.car_location[i][0] + 1))):
                if self.cells[j][y] == 1:
                    d2 = self.car_location[i][0] - j - self.car_length
                    flag = 0
                    break
            if flag:
                for j in range(self.lane_length - 1, int(self.car_location[i][0]), -1):
                    if self.cells[j][y] == 1:
                        d2 = self.car_location[i][0] - j - \
                            self.car_length + self.lane_length
                        flag = 0
                        break
            if flag:
                d2 = self.lane_length
            return d2

        def front(i):  # i为车辆序号，计算本车道前车距离
            flag = 1
            for j in range(int(self.car_location[i][0] + 1), self.lane_length):
                if self.cells[j][int(self.car_location[i][1])] == 1:
                    d1 = j - self.car_location[i][0] - self.car_length
                    flag = 0
                    break
            if flag:
                for j in range(0, int(self.car_location[i][0]+1)):
                    if self.cells[j][int(self.car_location[i][1])] == 1:
                        d1 = j - self.car_location[i][0] - \
                            self.car_length + self.lane_length
                        flag = 0
                        break
            return d1

        def other_front_v(i, y):
            flag = 1
            for a in range(int(self.rear[y]), int(self.head[y])+1):
                if self.car_location[i][0] < self.car_location[a][0]:
                    v3 = self.car_speed[a]
                    flag = 0
                    break

            if flag:
                v3 = self.car_speed[int(self.rear[y])]
                flag = 0
            if flag:
                v3 = 0
            return v3

        def other_back_v(i, y):  # i为车辆序号，y为潜在可变车道号，计算其他车道后车距离
            flag = 1
            for a in range(int(self.head[y]), int(self.rear[y])-1, -1):
                if self.car_location[i][0] >= self.car_location[a][0]:
                    v2 = self.car_speed[a]
                    flag = 0
                    break
            if flag:
                v2 = self.car_speed[int(self.head[y])]
                flag = 0
            if flag:
                v2 = 66
            return v2

        def front_v(i):
            if i != self.car_num-1:
                if self.car_location[i+1][1] == self.car_location[i][1]:
                    v1 = self.car_speed[i+1]
                else:
                    for a in range(self.car_num):
                        if self.car_location[a][1] == self.car_location[i][1]:
                            v1 = self.car_speed[a]
                            break
            if i == self.car_num-1:
                for a in range(self.car_num):
                    if self.car_location[a][1] == self.car_location[i][1]:
                        v1 = self.car_speed[a]
                        break
            return v1

        def count_cav(i):
            num = 0
            sspeed = 0
            cr = 600
            k = i
            if self.car_location[i][0]+cr < self.lane_length:  # 循环边界以内
                while(self.car_location[k][0] <= self.car_location[i][0]+cr) and (self.car_location[k][1] == self.car_location[i][1]) and (self.car_type[k] == 2):
                    num += 1
                    sspeed += self.car_speed[k]
                    if k < self.car_num - 1:
                        k += 1
                    else:
                        break

            elif self.car_location[i][0] + cr >= self.lane_length:  # 超过循环边界
                while (self.car_location[k][0] <= self.lane_length-1) and (self.car_location[k][1] == self.car_location[i][1]) and (self.car_type[k] == 2):
                    num += 1
                    sspeed += self.car_speed[k]
                    if k < self.car_num-1:
                        k += 1
                    else:
                        break
                while (self.car_location[k-1][1] == self.car_location[i][1]):
                    k -= 1
                while (self.car_location[k][0] <= self.car_location[i][0]+cr-self.lane_length) and (self.car_location[k][1] == self.car_location[i][1]) and (self.car_type[k] == 2):
                    num += 1
                    sspeed += self.car_speed[k]
                    k += 1

            if not num == 0:
                avspeed = sspeed / num
            else:
                avspeed = 0  # num为0时，设定平均速度为0
            return [num, avspeed]

        def count_other_cav(i, y):
            num = 0
            sspeed = 0
            cr = 600
            flag = 0
            if self.car_location[i][0] + cr < self.lane_length:
                for j in range(int(self.car_location[i][0]), int(self.car_location[i][0] + cr)):
                    for k in range(int(self.rear[y]), int(self.head[y])):
                        if self.car_location[k][0] == j and self.car_location[k][1] == y:
                            flag = 1
                            break
                    if flag:
                        break

                while (self.car_location[k][0] <= self.car_location[i][0] + cr) and (self.car_location[k][1] == y):
                    if self.car_type[k] == 2:
                        num += 1
                        sspeed += self.car_speed[k]
                    if k < self.car_num - 1:
                        k += 1
                    else:
                        break

            elif self.car_location[i][0] + cr >= self.lane_length:  # 超过循环边界
                for j in range(int(self.car_location[i][0]), self.lane_length):
                    for k in range(int(self.rear[y]), int(self.head[y])):
                        if self.car_location[k][0] == j and self.car_location[k][1] == y:
                            flag = 1
                            break
                    if flag:
                        break

                while (self.car_location[k][0] <= self.lane_length-1) and (self.car_location[k][1] == self.car_location[i][1]):
                    if self.car_type[k] == 2:
                        num += 1
                        sspeed += self.car_speed[k]
                    if k < self.car_num - 1:
                        k += 1
                    else:
                        break
                while (self.car_location[k-1][1] == y):
                    k -= 1
                while (self.car_location[k][0] <= self.car_location[i][0]+cr-self.lane_length) and (self.car_location[k][1] == y):
                    if self.car_type[k] == 2:
                        num += 1
                        sspeed += self.car_speed[k]
                    k += 1

            if not num == 0:
                avspeed = sspeed / num
            else:
                avspeed = 0  # num为0时，设定平均速度为0
            return [num, avspeed]

        def count_other_cav_num(i, y):
            num = 0
            sspeed = 0
            cr = 600
            flag = 0
            if self.car_location[i][0] + cr < self.lane_length:
                for j in range(int(self.car_location[i][0]), int(self.car_location[i][0] + cr)):
                    for k in range(int(self.car_num)):
                        if self.car_location[k][0] == j and self.car_location[k][1] == y:
                            flag = 1
                            break
                    if flag:
                        break

                while (self.car_location[k][0] <= self.car_location[i][0] + cr) and (self.car_location[k][1] == y) and (self.car_type[k] == 2):
                    num += 1
                    if k < self.car_num - 1:
                        k += 1
                    else:
                        break

            elif self.car_location[i][0] + cr >= self.lane_length:  # 超过循环边界
                for j in range(int(self.car_location[i][0]), self.lane_length):
                    for k in range(int(self.car_num)):
                        if self.car_location[k][0] == j and self.car_location[k][1] == y:
                            flag = 1
                            break
                    if flag:
                        break

                while (self.car_location[k][0] <= self.lane_length-1) and (self.car_location[k][1] == self.car_location[i][1]) and (self.car_type[k] == 2):
                    num += 1
                    if k < self.car_num - 1:
                        k += 1
                    else:
                        break
                while (self.car_location[k-1][1] == y):
                    k -= 1
                while (self.car_location[k][0] <= self.car_location[i][0]+cr-self.lane_length) and (self.car_location[k][1] == y) and (self.car_type[k] == 2):
                    num += 1
                    k += 1

            return num

        bmax = 3
        vmax = 66
        for i in range(int(self.car_num)):
            d_other_front = d_other_front2 = 0
            d_other_back = d_other_back2 = 0
            d_front = 0
            v_other_front = v_other_front2 = 0
            v_other_back = v_other_back2 = 0
            v_front = 0
            d_safe_front = d_safe_front2 = 0
            d_safe_back = d_safe_back2 = 0
            if self.car_type[i] == 1 and self.car_lc[i] == -1:  # 人工驾驶车
                if self.car_location[i][1] == 0:  # 最左侧车道（0号）
                    if self.lane_type[1] == 0 or self.lane_type[1] == 1:
                        d_front = front(i)
                        d_other_back = other_back(i, 1)
                        d_other_front = other_front(i, 1)
                        v_other_back = other_back_v(i, 1)
                        # 满足变道条件，以0.2概率变道
                        if d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = 1
                    else:
                        break

                elif self.car_location[i][1] == self.lane_num-1:  # 最右侧车道（3号）
                    if self.lane_type[self.lane_num-2] == 0 or self.lane_type[self.lane_num-2] == 1:
                        d_front = front(i)
                        d_other_back = other_back(i, self.lane_num-2)
                        d_other_front = other_front(i, self.lane_num-2)
                        v_other_back = other_back_v(i, self.lane_num-2)
                        if d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = self.lane_num-2
                    else:
                        break

                else:
                    if (self.lane_type[int(self.car_location[i][1])+1] == 0 or self.lane_type[int(self.car_location[i][1])+1] == 1) and \
                            (self.lane_type[int(self.car_location[i][1])-1] == 0 or self.lane_type[int(self.car_location[i][1])-1] == 1):  # 中间车道
                        d_front = front(i)
                        d_other_back = other_back(i, self.car_location[i][1]+1)
                        d_other_front = other_front(
                            i, self.car_location[i][1]+1)
                        d_other_back2 = other_back(
                            i, self.car_location[i][1]-1)
                        d_other_front2 = other_front(
                            i, self.car_location[i][1]-1)
                        v_other_back = other_back_v(i, self.car_location[i][1]+1)
                        v_other_back2 = other_back_v(i, self.car_location[i][1]-1)
                        if (d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0) and \
                                (d_front < min(self.car_speed[i] + 1, vmax) and d_other_front2 > d_front and d_other_back2 >= v_other_back2-self.car_speed[i] and d_other_back2>=0):
                            if d_other_front >= d_other_front2:  # 两条车道均可变时，选择车头间距更大的车道
                                if random.randint(0, 10) < 2:
                                    self.car_lc[i] = self.car_location[i][1]+1
                            else:
                                if random.randint(0, 10) < 2:
                                    self.car_lc[i] = self.car_location[i][1]-1

                        # 仅一条车道满足距离要求
                        elif d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = self.car_location[i][1]+1
                        elif d_front < min(self.car_speed[i] + 1, vmax) and d_other_front2 > d_front and d_other_back2 >= v_other_back2-self.car_speed[i] and d_other_back2>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = self.car_location[i][1]-1

                    # 仅一条车道满足专用道要求
                    elif (self.lane_type[int(self.car_location[i][1])+1] == 0 or self.lane_type[int(self.car_location[i][1])+1] == 1):
                        d_front = front(i)
                        d_other_back = other_back(
                            i, self.car_location[i][1] + 1)
                        d_other_front = other_front(
                            i, self.car_location[i][1] + 1)
                        v_other_back = other_back_v(i, self.car_location[i][1]+1)
                        if d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = self.car_location[i][1] + 1

                    elif (self.lane_type[int(self.car_location[i][1])-1] == 0 or self.lane_type[int(self.car_location[i][1])-1] == 1):
                        d_front = front(i)
                        d_other_back = other_back(
                            i, self.car_location[i][1] - 1)
                        d_other_front = other_front(
                            i, self.car_location[i][1] - 1)
                        v_other_back = other_back_v(i, self.car_location[i][1]-1)
                        if d_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                            if random.randint(0, 10) < 2:
                                self.car_lc[i] = self.car_location[i][1] - 1
                    else:
                        break

            if self.car_type[i] == 2 and self.car_lc[i] == -1:  # 智能网联车
                if self.car_location[i][1] == 0:  # 最左侧车道（0号）
                    if self.lane_type[1] == 0 or self.lane_type[1] == 2:
                        d_front = front(i)
                        d_other_back = other_back(i, 1)
                        d_other_front = other_front(i, 1)
                        v_front = front_v(i)
                        v_other_back = other_back_v(i, 1)
                        v_other_front = other_front_v(i, 1)
                        d_safe_front = (
                            self.car_speed[i]**2-v_other_front**2)/(2*bmax)
                        d_safe_back = (v_other_back**2 -
                                       self.car_speed[i]**2)/(2*bmax)
                        if d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front+v_front-v_other_front and d_other_front>=0 and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, 1) <= 4 or count_cav(i)[1] <= count_other_cav(i, 1)[1]):  # 满足变道条件，直接变道
                            self.car_lc[i] = 1
                    else:
                        break

                elif self.car_location[i][1] == self.lane_num-1:  # 最右侧车道（3号）
                    if self.lane_type[self.lane_num-2] == 0 or self.lane_type[self.lane_num-2] == 2:
                        d_front = front(i)
                        d_other_back = other_back(i, self.lane_num-2)
                        d_other_front = other_front(i, self.lane_num-2)
                        v_front = front_v(i)
                        v_other_back = other_back_v(i, self.lane_num-2)
                        v_other_front = other_front_v(i, self.lane_num-2)
                        d_safe_front = (
                            self.car_speed[i] ** 2 - v_other_front ** 2) / (2 * bmax)
                        d_safe_back = (v_other_back ** 2 -
                                       self.car_speed[i] ** 2) / (2 * bmax)
                        if d_front + v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front + v_front-v_other_front and d_other_front>=0\
                                and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.lane_num-2) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.lane_num-2)[1]):  # 满足变道条件，直接变道
                            self.car_lc[i] = self.lane_num-2
                    else:
                        break

                else:
                    if (self.lane_type[int(self.car_location[i][1])+1] == 0 or self.lane_type[int(self.car_location[i][1])+1] == 2) and \
                            (self.lane_type[int(self.car_location[i][1])-1] == 0 or self.lane_type[int(self.car_location[i][1])-1] == 2):
                        d_front = front(i)
                        d_other_back = other_back(i, self.car_location[i][1]+1)
                        d_other_front = other_front(
                            i, self.car_location[i][1]+1)
                        d_other_back2 = other_back(
                            i, self.car_location[i][1]-1)
                        d_other_front2 = other_front(
                            i, self.car_location[i][1]-1)
                        v_front = front(i)
                        v_other_back = other_back(i, self.car_location[i][1]+1)
                        v_other_front = other_front(
                            i, self.car_location[i][1]+1)
                        v_other_back2 = other_back(
                            i, self.car_location[i][1]-1)
                        v_other_front2 = other_front(
                            i, self.car_location[i][1]-1)
                        d_safe_front2 = (
                            self.car_speed[i] ** 2 - v_other_front2 ** 2) / (2 * bmax)
                        d_safe_back2 = (v_other_back2 ** 2 -
                                        self.car_speed[i] ** 2) / (2 * bmax)
                        if (d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front+v_front-v_other_front and d_other_front>=0 
                        and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0 and d_front+v_front < min(self.car_speed[i] + 1, vmax) 
                        and d_other_front2 > d_front+v_front-v_other_front2 and d_other_front2>=0 \
                            and d_other_back2 >= v_other_back2 - self.car_speed[i] and d_other_back2==0):
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.car_location[i][1]-1) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.car_location[i][1]-1)[1])):
                            # 两条车道均可变时，选择CR中智能网联车更多的车道
                            if count_other_cav_num(i, self.car_location[i][1]+1) >= count_other_cav_num(i, self.car_location[i][1]-1):
                                self.car_lc[i] = self.car_location[i][1]+1
                            else:
                                self.car_lc[i] = self.car_location[i][1]-1
                        elif d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front+v_front-v_other_front and d_other_front>=0\
                                and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.car_location[i][1]+1) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.car_location[i][1]+1)[1]):
                            self.car_lc[i] = self.car_location[i][1]+1

                        elif d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front2 > d_front+v_front-v_other_front2 and d_other_front2>=0 \
                                and d_other_back2 >= v_other_back2-self.car_speed[i] and d_other_back2>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.car_location[i][1]-1) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.car_location[i][1]-1)[1]):
                            self.car_lc[i] = self.car_location[i][1]-1

                    elif (self.lane_type[int(self.car_location[i][1])+1] == 0 or self.lane_type[int(self.car_location[i][1])+1] == 2):
                        d_front = front(i)
                        d_other_back = other_back(
                            i, self.car_location[i][1] + 1)
                        d_other_front = other_front(
                            i, self.car_location[i][1] + 1)
                        v_front = front(i)
                        v_other_back = other_back(i, self.car_location[i][1]+1)
                        v_other_front = other_front(
                            i, self.car_location[i][1]+1)
                        d_safe_front = (
                            self.car_speed[i] ** 2 - v_other_front ** 2) / (2 * bmax)
                        d_safe_back = (v_other_back ** 2 -
                                       self.car_speed[i] ** 2) / (2 * bmax)
                        if d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front+v_front-v_other_front and d_other_front>=0 \
                                and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.car_location[i][1]+1) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.car_location[i][1]+1)[1]):
                            self.car_lc[i] = self.car_location[i][1] + 1

                    elif (self.lane_type[int(self.car_location[i][1])-1] == 0 or self.lane_type[int(self.car_location[i][1])-1] == 2):
                        d_front = front(i)
                        d_other_back = other_back(
                            i, self.car_location[i][1] - 1)
                        d_other_front = other_front(
                            i, self.car_location[i][1] - 1)
                        v_front = front(i)
                        v_other_back = other_back(i, self.car_location[i][1]-1)
                        v_other_front = other_front(
                            i, self.car_location[i][1]-1)
                        d_safe_front = (
                            self.car_speed[i] ** 2 - v_other_front ** 2) / (2 * bmax)
                        d_safe_back = (v_other_back ** 2 -
                                       self.car_speed[i] ** 2) / (2 * bmax)
                        if d_front+v_front < min(self.car_speed[i] + 1, vmax) and d_other_front > d_front+v_front-v_other_front and d_other_front>=0 \
                                and d_other_back >= v_other_back-self.car_speed[i] and d_other_back>=0:
                                #and (count_cav(i)[0] <= count_other_cav_num(i, self.car_location[i][1]-1) <= 4 or count_cav(i)[1] <= count_other_cav(i, self.car_location[i][1]-1)[1]):
                            self.car_lc[i] = self.car_location[i][1] - 1

                    else:
                        break

            if self.car_lc[i] != -1 and self.car_location[i][1] != self.car_lc[i]: # 开始换道 car_lc非-1的车辆换道
                self.cells[int(self.car_location[i][0])
                           ][int(self.car_location[i][1])] = 0
                self.cells[int(self.car_location[i][0])
                           ][int(self.car_lc[i])] = 1

                if self.timer>=2000: #换道次数统计
                    self.change_in_nums[int(self.car_lc[i])]+=1
                    self.change_out_nums[int(self.car_location[i][1])]+=1
                # print(self.car_location[i][1],self.car_lc[i],self.car_type[i])
                self.car_location[i][1] = self.car_lc[i]
                #print(self.car_type[i])
                flag = 1
                for m in range(int(self.car_num)):
                    if self.car_location[i][0]+self.car_location[i][1]*self.lane_length < self.car_location[m][0]+self.car_location[m][1]*self.lane_length:
                        flag = 0
                        break
                if flag:
                    m = self.car_num
                #print(i,m)
                if m > i:     #向右变道
                    tmp = self.car_location[i]
                    tmp1 = self.car_speed[i]
                    tmp2 = self.car_type[i]
                    tmp3 = self.car_lc[i]
                    for k in range(i, m-1):
                        self.car_location[k] = self.car_location[k + 1]
                        self.car_speed[k] = self.car_speed[k + 1]
                        self.car_type[k] = self.car_type[k + 1]
                        self.car_lc[k] = self.car_lc[k+1]
                    self.car_location[m-1] = tmp
                    self.car_speed[m-1] = tmp1
                    self.car_type[m-1] = tmp2
                    self.car_lc[m-1] = tmp3
                    i -= 1
                elif m < i:   #向左变道
                    tmp = self.car_location[i]
                    tmp1 = self.car_speed[i]
                    tmp2 = self.car_type[i]
                    tmp3 = self.car_lc[i]
                    for k in range(i, m, -1):
                        self.car_location[k] = self.car_location[k - 1]
                        self.car_speed[k] = self.car_speed[k - 1]
                        self.car_type[k] = self.car_type[k - 1]
                        self.car_lc[k] = self.car_lc[k - 1]
                    self.car_location[m] = tmp
                    self.car_speed[m] = tmp1
                    self.car_type[m] = tmp2
                    self.car_lc[m] = tmp3


        for i in range(self.car_num):  # 换道标号重置-1
            self.car_lc[i] = -1

    def print_road(self):  # 输出道路状态，该位置CAV为&该位置MV为*，无车为_
        for i in range(self.lane_num):
            for j in range(self.lane_length):
                if self.cells[j][i] == 1:
                    for m in range(int(self.car_num)):
                        if int(self.car_location[m][0]) == j and int(self.car_location[m][1]) == i:
                            if self.car_type[m] == 2:
                                print("&", end='')
                                break
                            elif self.car_type[m] == 1:
                                print("*", end='')
                                break
                elif self.cells[j][i] == 0:
                    print("_", end='')
            print("\n")

    def print_cars(self):  # 输出车辆状态
        print(self.car_location)
        print(self.car_speed)
        # print(self.car_type)


    def Iter(self, n):
        workbook = xlsxwriter.Workbook(
             self.elusive_lane + '_'+str(self.car_num)+'_'+str(self.cav_proportion)+'_'+str(self.num)+'.xlsx')
        #worksheet = workbook.add_worksheet('Sheet 2')
        for _ in range(n):
            self.lane_change()
            self.update_state()
            if 2000<=self.timer<5600:
                for i in range(4):
                    self.veh_num[i][self.timer-2000]=self.head[i]-self.rear[i]+1
                #worksheet.write(self.timer-1999,1,self.change_in_nums[0])
                #worksheet.write(self.timer-1999,2,self.change_in_nums[1])
                #worksheet.write(self.timer-1999,3,self.change_in_nums[2])
                #worksheet.write(self.timer-1999,4,self.change_in_nums[3])
            # self.print_road()
            # self.print_cars()
            #print(_)

        worksheet = workbook.add_worksheet('Sheet 1')
        worksheet.write(0, 2, '小时流量')
        worksheet.write(0, 3, '平均速度')
        worksheet.write(0, 4, '换进次数')
        worksheet.write(0, 5, '换出次数')
        worksheet.write(0, 6, '平均车辆密度')
        #worksheet.write(0, 7, 'MV车辆数')
        for i in range(self.lane_num):
            worksheet.write(i+1, 0, '车道'+str(i+1))
            if self.lane_type[i] == 0:
                worksheet.write(i+1, 1, "普通车道")
            elif self.lane_type[i] == 1:
                worksheet.write(i+1, 1, "人工驾驶专用道")
            elif self.lane_type[i] == 2:
                worksheet.write(i+1, 1, "智能网联专用道")
            worksheet.write(i+1, 2, self.flux[i])
            worksheet.write(i+1, 3, np.mean(self.lane_avspeed[i]))
            worksheet.write(i+1, 4, self.change_in_nums[i])
            worksheet.write(i+1, 5, self.change_out_nums[i])
            worksheet.write(i+1, 6, np.mean(self.veh_num[i])/3)
            #worksheet.write(i+1, 7, self.mv_nums[i]/3600)
        
        worksheet.write(0, 10, 'MV平均速度')
        worksheet.write(1, 10, 'CAV平均速度')
        worksheet.write(0, 11, np.mean(self.MV_avspeed))
        worksheet.write(1, 11, np.mean(self.CAV_avspeed))
        workbook.close()


car_num = 0
cav_proportion = 0
car_length = 15
elusive_lane = ''
data = xlrd.open_workbook("参数设置.xlsx")
sheet1 = data.sheets()[0]
i=1
while int(sheet1.cell(i, 0).value):
    car_num = int(sheet1.cell(i, 0).value)
    cav_proportion = float(sheet1.cell(i, 1).value)
    elusive_lane = str(sheet1.cell(i, 2).value)
    run_time = int(sheet1.cell(i, 3).value)
    print(car_num, cav_proportion, elusive_lane)
    i+=1
    for j in range(run_time):
        # (car_num,car_length,cells_shape,cav_proportion,elusive_lane)每格0.5m
        automata = Automata(car_num, car_length, (6000, 4),
                            cav_proportion, elusive_lane, j+1)
        print("正在进行第"+str(j+1)+"次模拟")
        automata.car_distribute()
        automata.Iter(5600)
        print("第"+str(j+1)+"次模拟完成")
