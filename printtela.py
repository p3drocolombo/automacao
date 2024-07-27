import cv2
import numpy as np
from selenium import webdriver

# Lendo a imagem usando OpenCV
image = cv2.imread("pagina_screenshot.png")

# Exibindo a imagem para que o usuário possa selecionar a área de corte
clone = image.copy()
selected_points = []

def click_and_select(event, x, y, flags, param):
    global selected_points

    if event == cv2.EVENT_LBUTTONDOWN:
        selected_points.append((x, y))
        cv2.circle(clone, (x, y), 3, (0, 255, 0), -1)
        cv2.imshow("Selecionar área de corte", clone)

        if len(selected_points) == 4:
            cv2.destroyAllWindows()

cv2.imshow("Selecionar área de corte", clone)
cv2.setMouseCallback("Selecionar área de corte", click_and_select)
cv2.waitKey(0)

# Calculando coordenadas da área selecionada
x, y, width, height = cv2.boundingRect(np.array(selected_points))
print("Coordenadas da área selecionada (x, y, largura, altura):", x, y, width, height)

# Salvar as coordenadas em um arquivo
with open('coordenadas.txt', 'w') as f:
    f.write(f"{x},{y},{width},{height}")

# Cortando a imagem
cropped_image = image[y:y+height, x:x+width]

# Mostrando a imagem cortada
cv2.imshow("Imagem cortada", cropped_image)
cv2.waitKey(0)
cv2.destroyAllWindows()
