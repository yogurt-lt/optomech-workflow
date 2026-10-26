"""Export a geometric travel law, not calibrated optical focus distances."""
import argparse,csv,math

def samples(angle_start,angle_end,travel_start,travel_end,count,law,input_ratio=None):
    values=[angle_start,angle_end,travel_start,travel_end]
    if input_ratio is not None:values.append(input_ratio)
    if not all(math.isfinite(v) for v in values):raise ValueError('Endpoints and ratio must be finite.')
    if angle_end<=angle_start or not 2<=count<=10001:raise ValueError('Require increasing angles and 2–10001 samples.')
    if law not in ('linear','smoothstep'):raise ValueError('Unknown geometric law.')
    for i in range(count):
        u=i/(count-1);f=u if law=='linear' else u*u*(3-2*u)
        angle=angle_start+(angle_end-angle_start)*u
        yield [angle,travel_start+(travel_end-travel_start)*f]+([] if input_ratio is None else [angle*input_ratio])

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    for name in ['angle-start','angle-end','travel-start','travel-end']:p.add_argument('--'+name,type=float,required=True)
    p.add_argument('--samples',type=int,default=21);p.add_argument('--law',choices=['linear','smoothstep'],required=True);p.add_argument('--input-ratio',type=float);p.add_argument('--output',required=True);a=p.parse_args()
    rows=list(samples(a.angle_start,a.angle_end,a.travel_start,a.travel_end,a.samples,a.law,a.input_ratio))
    with open(a.output,'w',newline='',encoding='utf8') as f:
        w=csv.writer(f);w.writerow(['angle_deg','travel_mm']+([] if a.input_ratio is None else ['input_angle_deg']));w.writerows(rows)
