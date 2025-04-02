import pygame
import sys
import math
from pygame.locals import *

# Initialize pygame
pygame.init()

# Constants
WIDTH, HEIGHT = 800, 600
FPS = 60
BACKGROUND_COLOR = (0, 0, 0)

# Physics constants
GRAVITY = 0.5
FRICTION = 0.98
ELASTICITY = 1.0  # 增加这个值可以让小球更有弹性（原值为0.8）
GROUND_ATTRACTION = 0.2

# Triangle properties
TRIANGLE_SIZE = 250
TRIANGLE_COLOR = (255, 255, 255)
TRIANGLE_ROTATION_SPEED = 0.5  # degrees per frame

# Ball properties
BALL_RADIUS = 15
BALL_COLOR = (255, 0, 0)
BALL_MASS = 10

# Create the screen
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Rotating Triangle with Ball Physics")
clock = pygame.time.Clock()

class Triangle:
    def __init__(self, size, center_x, center_y):
        self.size = size
        self.center_x = center_x
        self.center_y = center_y
        self.angle = 0
        self.vertices = self.calculate_vertices()
        
    def calculate_vertices(self):
        # Calculate the vertices of the triangle based on the center and current angle
        vertices = []
        for i in range(3):
            angle_rad = math.radians(self.angle + i * 120)
            x = self.center_x + self.size * math.cos(angle_rad)
            y = self.center_y + self.size * math.sin(angle_rad)
            vertices.append((x, y))
        return vertices
    
    def rotate(self, speed):
        self.angle += speed
        self.vertices = self.calculate_vertices()
    
    def draw(self, surface):
        pygame.draw.polygon(surface, TRIANGLE_COLOR, self.vertices, 2)
    
    def get_edges(self):
        # Return the edges as pairs of points
        edges = []
        for i in range(3):
            edges.append((self.vertices[i], self.vertices[(i + 1) % 3]))
        return edges

class Ball:
    def __init__(self, x, y, radius, color, mass):
        self.x = x
        self.y = y
        self.vx = 0
        self.vy = 0
        self.radius = radius
        self.color = color
        self.mass = mass
    
    def update(self, triangle):
        # Apply gravity
        self.vy += GRAVITY
        
        # Apply ground attraction (pulls toward the bottom of the screen)
        self.vy += GROUND_ATTRACTION * (HEIGHT - self.y) / HEIGHT
        
        # Update position
        self.x += self.vx
        self.y += self.vy
        
        # Apply friction
        self.vx *= FRICTION
        self.vy *= FRICTION
        
        # Check for collision with triangle edges
        self.check_collision(triangle)
    
    def check_collision(self, triangle):
        edges = triangle.get_edges()
        for edge in edges:
            # Check collision with each edge
            self.handle_edge_collision(edge)
        
        # Ensure the ball is inside the triangle
        if not self.is_inside_triangle(triangle.vertices):
            # If not inside, find the closest edge and push the ball inside
            self.force_inside_triangle(triangle.vertices, edges)
    
    def handle_edge_collision(self, edge):
        # Line segment defined by p1 and p2
        p1, p2 = edge
        
        # Vector from p1 to p2
        edge_vector = (p2[0] - p1[0], p2[1] - p1[1])
        edge_length = math.sqrt(edge_vector[0]**2 + edge_vector[1]**2)
        
        # Normalized edge vector
        if edge_length == 0:
            return  # Avoid division by zero
        edge_unit = (edge_vector[0] / edge_length, edge_vector[1] / edge_length)
        
        # Vector from p1 to ball center
        to_ball = (self.x - p1[0], self.y - p1[1])
        
        # Project to_ball onto edge_unit to get the closest point on the line
        projection = to_ball[0] * edge_unit[0] + to_ball[1] * edge_unit[1]
        
        # Clamp projection to the line segment
        projection = max(0, min(edge_length, projection))
        
        # Calculate closest point on the line segment
        closest_point = (
            p1[0] + projection * edge_unit[0],
            p1[1] + projection * edge_unit[1]
        )
        
        # Calculate distance between ball and closest point
        dx = self.x - closest_point[0]
        dy = self.y - closest_point[1]
        distance = math.sqrt(dx*dx + dy*dy)
        
        # Check if collision occurs
        if distance <= self.radius:
            # Calculate overlap
            overlap = self.radius - distance
            
            # Normal vector from edge to ball (normalized)
            if distance == 0:  # Avoid division by zero
                nx, ny = 0, -1  # Default normal
            else:
                nx, ny = dx / distance, dy / distance
                
            # Move ball away from edge based on overlap
            self.x += nx * overlap
            self.y += ny * overlap
            
            # Calculate reflection vector
            dot_product = self.vx * nx + self.vy * ny
            self.vx = (self.vx - 2 * dot_product * nx) * ELASTICITY
            self.vy = (self.vy - 2 * dot_product * ny) * ELASTICITY
    
    def is_inside_triangle(self, vertices):
        # Check if the ball is inside the triangle using barycentric coordinates
        def sign(p1, p2, p3):
            return (p1[0] - p3[0]) * (p2[1] - p3[1]) - (p2[0] - p3[0]) * (p1[1] - p3[1])
            
        d1 = sign((self.x, self.y), vertices[0], vertices[1])
        d2 = sign((self.x, self.y), vertices[1], vertices[2])
        d3 = sign((self.x, self.y), vertices[2], vertices[0])
        
        has_neg = (d1 < 0) or (d2 < 0) or (d3 < 0)
        has_pos = (d1 > 0) or (d2 > 0) or (d3 > 0)
        
        # If all signs are the same, the point is inside the triangle
        return not (has_neg and has_pos)
    
    def force_inside_triangle(self, vertices, edges):
        # Find the closest edge and push the ball inside
        min_distance = float('inf')
        closest_edge = None
        closest_point = None
        
        for edge in edges:
            # Line segment defined by p1 and p2
            p1, p2 = edge
            
            # Vector from p1 to p2
            edge_vector = (p2[0] - p1[0], p2[1] - p1[1])
            edge_length = math.sqrt(edge_vector[0]**2 + edge_vector[1]**2)
            
            # Normalized edge vector
            if edge_length == 0:
                continue  # Avoid division by zero
            edge_unit = (edge_vector[0] / edge_length, edge_vector[1] / edge_length)
            
            # Vector from p1 to ball center
            to_ball = (self.x - p1[0], self.y - p1[1])
            
            # Project to_ball onto edge_unit to get the closest point on the line
            projection = to_ball[0] * edge_unit[0] + to_ball[1] * edge_unit[1]
            
            # Clamp projection to the line segment
            projection = max(0, min(edge_length, projection))
            
            # Calculate closest point on the line segment
            point = (
                p1[0] + projection * edge_unit[0],
                p1[1] + projection * edge_unit[1]
            )
            
            # Calculate distance between ball and closest point
            dx = self.x - point[0]
            dy = self.y - point[1]
            distance = math.sqrt(dx*dx + dy*dy)
            
            if distance < min_distance:
                min_distance = distance
                closest_edge = edge
                closest_point = point
        
        if closest_edge:
            # Calculate direction toward inside of triangle
            p1, p2 = closest_edge
            # Get third vertex to determine inside direction
            third_vertex = None
            for v in vertices:
                if v != p1 and v != p2:
                    third_vertex = v
                    break
            
            # Calculate the normal vector to the edge
            edge_vector = (p2[0] - p1[0], p2[1] - p1[1])
            # Perpendicular vector (rotated 90 degrees)
            normal = (-edge_vector[1], edge_vector[0])
            
            # Normalize the normal
            norm_length = math.sqrt(normal[0]**2 + normal[1]**2)
            if norm_length > 0:
                normal = (normal[0] / norm_length, normal[1] / norm_length)
            
            # Check if normal points inside the triangle
            test_vector = (third_vertex[0] - closest_point[0], 
                           third_vertex[1] - closest_point[1])
            dot_product = normal[0] * test_vector[0] + normal[1] * test_vector[1]
            
            # Flip normal if needed
            if dot_product < 0:
                normal = (-normal[0], -normal[1])
            
            # Move ball inside along the normal
            move_distance = self.radius + 1  # Move a bit more than radius
            self.x = closest_point[0] + normal[0] * move_distance
            self.y = closest_point[1] + normal[1] * move_distance
    
    def draw(self, surface):
        pygame.draw.circle(surface, self.color, (int(self.x), int(self.y)), self.radius)

# Create objects
triangle = Triangle(TRIANGLE_SIZE, WIDTH // 2, HEIGHT // 2)
ball = Ball(WIDTH // 2, HEIGHT // 2, BALL_RADIUS, BALL_COLOR, BALL_MASS)

# Main game loop
running = True
while running:
    for event in pygame.event.get():
        if event.type == QUIT:
            running = False
    
    # Fill the screen
    screen.fill(BACKGROUND_COLOR)
    
    # Update triangle
    triangle.rotate(TRIANGLE_ROTATION_SPEED)
    
    # Update ball position and handle collision
    ball.update(triangle)
    
    # Draw objects
    triangle.draw(screen)
    ball.draw(screen)
    
    # Update the display
    pygame.display.flip()
    clock.tick(FPS)

pygame.quit()
sys.exit()
