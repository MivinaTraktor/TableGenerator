from math import sin, cos, atan, sqrt, degrees, radians, pi
import pandas as pd
import openpyxl

deg = 6400
dist_step = 25  # meters
alt_dif = (50, 100, 200, 300)  # meters
side_step = (20,)  # meters
gun_limits = (800, 1460)  # float for degrees, int for mils
# m6: 800, 1500
g = 9.81


def main():
    df = calc_table([265.409], 0.4432, 0.583, 0.6672, 0.7656, 0.8377, 0.9014, 1.0001)
    # m6: [153], 0.58, 0.72, 0.85, 1
    # m120: [310], 0.4432, 0.583, 0.6672, 0.7656, 0.8377, 0.9014, 1.0001
    # 2b11 [265.409], 0.4432, 0.583, 0.6672, 0.7656, 0.8377, 0.9014, 1.0001
    print(df)
    df.to_excel('tables.xlsx', sheet_name='2b11')


def calc_table(init_speeds, *multipliers):
    low, hi = (i if isinstance(i, float) else i / deg * 360 for i in gun_limits)
    if low < 45:
        low = 45
    if hi > 89:
        hi = 89
    if len(init_speeds) == 1:
        init_speeds = [init_speeds[0] * i for i in multipliers]
    table = {}
    charge = 0
    for speed in init_speeds:
        table_charge = {}
        min = dist_by_angle(speed, hi) // dist_step * dist_step + dist_step
        max = dist_by_angle(speed, low) // dist_step * dist_step
        dist = min
        while dist <= max:
            table_charge[f'{dist}m'] = calc_line(speed, dist)
            dist += dist_step
        table[f'charge {charge}'] = table_charge
        charge += 1
    df = pd.DataFrame.from_dict({(i, j): table[i][j]
                                 for i in table.keys()
                                 for j in table[i].keys()},
                                orient='index')
    return df


def calc_line(v, dist):
    el = elev_formula(v, dist)
    el_deg = degrees(el) / 360 * deg
    if el is None:
        return None
    delta = [
        degrees(elev_formula(v, dist, alt)) / 360 * deg - el_deg if elev_formula(v, dist, alt) is not None else None for
        alt in alt_dif]
    tof = 2 * v * sin(el) / g
    one_deg = m_per_deg(dist)
    a = {'elevation': el_deg}
    a.update({f'^{value}m': delta[i] for i, value in enumerate(alt_dif)})
    # a.update({'-> per 1deg': one_deg})
    # a.update({f'-> for {i}m': i / one_deg for i in side_step})
    a.update({'TOF': tof})
    return a


def elev_formula(v, dist, alt_dif=None):
    try:
        if alt_dif is None:
            rad = atan((v ** 2 + sqrt(v ** 4 - g * (g * dist ** 2))) / (g * dist))
        else:
            rad = atan((v ** 2 + sqrt(v ** 4 - g * ((g * dist ** 2) + (2 * alt_dif * v ** 2)))) / (g * dist))
    except Exception:
        return None
    return rad


def dist_by_angle(v, angle):
    return 2 * v ** 2 / g * cos(radians(angle)) * sin(radians(angle))


def m_per_deg(dist):
    return 2 * pi * dist / deg


if __name__ == '__main__':
    main()
