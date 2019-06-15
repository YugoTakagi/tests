import numpy as np
import matplotlib.pyplot as plt
import random
import matplotlib.patches as patches


def main():
    n = 100
    m = 0.17
    g = 9.81

    xx=[]
    yy=[]
    zz=[]
    S = []
    time=[]
    xi = (45.0/180.0) *np.pi

    # v = 7.0
    v = 7.5
    vxO = v *np.cos(xi)
    vzO = v *np.sin(xi)
    h = 0.3

    mu = v
    sigma = 0.1#±0.3m/sズレてもいい感じ？
    vel = [random.gauss(mu, sigma) for i in range(n)]

    for v in vel:
        vxO = v *np.cos(xi)
        vzO = v *np.sin(xi)
        time = (-vzO -np.sqrt(vzO**2 -4*(-g/2)*h))/(2*(-g/2))#解の公式   t(vyO, h)
        for t in np.arange(0,time, 0.01):
            xx.append(vxO *t)
            zz.append(vzO *t -(g* t**2)/2 +h)
        S.append(xx[len(xx)-1])
        # plt.plot(xx, zz, ".")
    # plt.plot(time, vv)
    # plt.show()
    # print(S)

    fig = plt.figure()
    ax = plt.axes()
    r = patches.Rectangle(xy=(-0.75/2, -0.75/2), width=3.5, height=7.0, ec='#000000', fill=False)
    ax.add_patch(r)
    r = patches.Rectangle(xy=(-0.75/2, -0.75/2), width=3.5, height=7.0/2, ec='#000000', fill=False)
    ax.add_patch(r)
    r = patches.Rectangle(xy=(3.0-0.75/2, 5.0-0.75/2), width=1.0, height=1.0, ec='#000000', fill=False)
    ax.add_patch(r)
    # plt.axis('scaled')

    # xRef = 3.5/2 -0.75/2
    # yRef = 7.0 -0.75/2
    xRef = 3.0 -0.75/2
    yRef = 5.0 -0.75/2
    xiRef = np.arctan(yRef/xRef)
    # print(xiRef)

    mu = np.pi/2 -xiRef
    # print(mu, mu*180/np.pi)
    sigma = 1.0/180*np.pi#±5度くらい？
    xita = [random.gauss(mu, sigma) for i in range(n)]
    xxXi=[]
    yyXi=[]
    for i, e in enumerate(xita):
        xxXi.append(S[i] *np.sin(e))
        yyXi.append(S[i] *np.cos(e))
        plt.plot(xxXi, yyXi, ".")
    # plt.plot(time, vv)
    plt.show()



if __name__ == "__main__":
    main()
