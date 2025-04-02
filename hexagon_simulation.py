import pygame, math, sys

# 初始化
pygame.init()
width, height = 800, 600
screen = pygame.display.set_mode((width, height))
clock = pygame.time.Clock()

# 六边形参数
hex_radius = 250
center = (width//2, height//2)
angle_offset = math.radians(-30)

# 修改：get_hexagon_vertices 增加 angle_off 参数
def get_hexagon_vertices(center, radius, angle_off):
    cx, cy = center
    vertices = []
    for i in range(6):
        angle = angle_off + math.radians(60 * i)
        x = cx + radius * math.cos(angle)
        y = cy + radius * math.sin(angle)
        vertices.append((x, y))
    return vertices

# 初始化旋转变量
current_angle_offset = angle_offset
rotation_speed = math.radians(30)  # 30°/秒

hex_vertices = get_hexagon_vertices(center, hex_radius, current_angle_offset)

# 新增：判断点是否在多边形内，采用射线法
def point_in_polygon(point, polygon):
    x, y = point
    inside = False
    n = len(polygon)
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi + 1e-10) + xi):
            inside = not inside
        j = i
    return inside

# 计算直线到点的距离及最近点（直线段AB到点 P）
def point_line_distance(P, A, B):
    # 计算投影
    AP = (P[0]-A[0], P[1]-A[1])
    AB = (B[0]-A[0], B[1]-A[1])
    ab2 = AB[0]**2 + AB[1]**2
    if ab2 == 0:
        return math.hypot(P[0]-A[0], P[1]-A[1]), A
    t = max(0, min(1, (AP[0]*AB[0] + AP[1]*AB[1])/ab2))
    closest = (A[0]+AB[0]*t, A[1]+AB[1]*t)
    dist = math.hypot(P[0]-closest[0], P[1]-closest[1])
    return dist, closest

class Ball:
    def __init__(self, pos, vel, radius, color, mass=1):
        self.pos = list(pos)
        self.vel = list(vel)
        self.radius = radius
        self.color = color
        self.mass = mass

    def update(self, dt):
        self.pos[0] += self.vel[0] * dt
        self.pos[1] += self.vel[1] * dt

    def draw(self, surface):
        pygame.draw.circle(surface, self.color, (int(self.pos[0]), int(self.pos[1])), self.radius)

    def wall_collision(self, vertices):
        # 针对六边形的每条边检测碰撞
        for i in range(len(vertices)):
            A = vertices[i]
            B = vertices[(i+1) % len(vertices)]
            dist, closest = point_line_distance(self.pos, A, B)
            if dist < self.radius:
                # 计算边外法向量
                edge = (B[0]-A[0], B[1]-A[1])
                # 取边的垂直向量
                normal = (-edge[1], edge[0])
                # 修正法向量方向：应当指向小球外部（与六边形中心相反）
                mid = ((A[0]+B[0])/2, (A[1]+B[1])/2)
                dir_vec = (self.pos[0]-mid[0], self.pos[1]-mid[1])
                if (normal[0]*dir_vec[0] + normal[1]*dir_vec[1]) < 0:
                    normal = (-normal[0], -normal[1])
                # 归一化 normal
                n_len = math.hypot(normal[0], normal[1])
                if n_len == 0:
                    continue
                normal = (normal[0]/n_len, normal[1]/n_len)
                # 反射速度： v' = v - 2*(v·n)*n
                v_dot_n = self.vel[0]*normal[0] + self.vel[1]*normal[1]
                if v_dot_n < 0:  # 只有在朝墙方向时才反弹
                    self.vel[0] -= 2 * v_dot_n * normal[0]
                    self.vel[1] -= 2 * v_dot_n * normal[1]
                    # 调整位置以防止粘连
                    overlap = self.radius - dist
                    self.pos[0] += normal[0]*overlap
                    self.pos[1] += normal[1]*overlap

def ball_collision(ball1, ball2):
    dx = ball2.pos[0] - ball1.pos[0]
    dy = ball2.pos[1] - ball1.pos[1]
    dist = math.hypot(dx, dy)
    if dist < ball1.radius + ball2.radius and dist != 0:
        # 单位法向量
        nx = dx/dist
        ny = dy/dist
        # 速度在法线方向的分量
        dvx = ball1.vel[0] - ball2.vel[0]
        dvy = ball1.vel[1] - ball2.vel[1]
        impact_speed = dvx*nx + dvy*ny
        if impact_speed > 0:
            return
        # 计算冲量 (完全弹性碰撞)
        impulse = 2 * impact_speed / (ball1.mass + ball2.mass)
        ball1.vel[0] -= impulse * ball2.mass * nx
        ball1.vel[1] -= impulse * ball2.mass * ny
        ball2.vel[0] += impulse * ball1.mass * nx
        ball2.vel[1] += impulse * ball1.mass * ny
        # 简单分离，使两个球不重叠
        overlap = ball1.radius + ball2.radius - dist
        ball1.pos[0] -= (overlap/2) * nx
        ball1.pos[1] -= (overlap/2) * ny
        ball2.pos[0] += (overlap/2) * nx
        ball2.pos[1] += (overlap/2) * ny

# 修改：增加小球到 5 个
red_ball = Ball(pos=(center[0]-100, center[1]), vel=(150, 120), radius=15, color=(255,0,0))
blue_ball = Ball(pos=(center[0]+100, center[1]), vel=(-130, -100), radius=15, color=(0,0,255))
green_ball = Ball(pos=(center[0], center[1]-100), vel=(100, 150), radius=15, color=(0,255,0))
orange_ball = Ball(pos=(center[0]-80, center[1]+80), vel=(120, -140), radius=15, color=(255,165,0))
purple_ball = Ball(pos=(center[0]+80, center[1]+80), vel=(-140, 130), radius=15, color=(128,0,128))
balls = [red_ball, blue_ball, green_ball, orange_ball, purple_ball]

running = True
while running:
    dt = clock.tick(60) / 1000.0  # 秒
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False

    # 更新六边形旋转（顺时针旋转）
    current_angle_offset -= rotation_speed * dt
    hex_vertices = get_hexagon_vertices(center, hex_radius, current_angle_offset)

    # 更新小球位置并检查墙面碰撞
    for ball in balls:
        ball.update(dt)
        ball.wall_collision(hex_vertices)
        # 新增：如果小球不在六边形内，则修正其位置和速度
        if not point_in_polygon(ball.pos, hex_vertices):
            min_dist = float('inf')
            best_normal = None
            best_point = None
            for i in range(len(hex_vertices)):
                A = hex_vertices[i]
                B = hex_vertices[(i+1)%len(hex_vertices)]
                d, close = point_line_distance(ball.pos, A, B)
                if d < min_dist:
                    min_dist = d
                    edge = (B[0]-A[0], B[1]-A[1])
                    normal = (-edge[1], edge[0])
                    mid = ((A[0]+B[0])/2, (A[1]+B[1])/2)
                    dir_vec = (ball.pos[0]-mid[0], ball.pos[1]-mid[1])
                    if (normal[0]*dir_vec[0] + normal[1]*dir_vec[1]) < 0:
                        normal = (-normal[0], -normal[1])
                    n_len = math.hypot(normal[0], normal[1])
                    if n_len != 0:
                        best_normal = (normal[0]/n_len, normal[1]/n_len)
                    else:
                        best_normal = (0,0)
                    best_point = close
            if best_normal and best_point:
                # 将小球置于边界内（刚好在小球半径距离处）
                ball.pos[0] = best_point[0] + best_normal[0]*ball.radius
                ball.pos[1] = best_point[1] + best_normal[1]*ball.radius
                # 反射速度
                v_dot_n = ball.vel[0]*best_normal[0] + ball.vel[1]*best_normal[1]
                if v_dot_n < 0:
                    ball.vel[0] -= 2*v_dot_n*best_normal[0]
                    ball.vel[1] -= 2*v_dot_n*best_normal[1]

    # 小球之间碰撞
    for i in range(len(balls)):
        for j in range(i+1, len(balls)):
            ball_collision(balls[i], balls[j])

    # 绘制
    screen.fill((255,255,255))
    # 绘制六边形
    pygame.draw.polygon(screen, (0,0,0), hex_vertices, 3)
    for ball in balls:
        ball.draw(screen)
    pygame.display.flip()

pygame.quit()
sys.exit()
