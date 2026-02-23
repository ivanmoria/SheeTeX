import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from PIL import Image

# --- PARÂMETROS ---
fs = 44100  # taxa de amostragem
duration = 10  # segundos
t = np.linspace(0, duration, int(fs * duration))

# Frequências em Hz
f1 = 294  # D4
f2 = 392  # G4
f3 = 494  # B4

# Funções seno
wave1 = np.sin(2 * np.pi * f1 * t)
wave2 = np.sin(2 * np.pi * f2 * t)
wave3 = np.sin(2 * np.pi * f3 * t)

# Soma das ondas
wave_sum = wave1 + wave2 + wave3

# --- ANIMAÇÃO EM GIF ---
fig, ax = plt.subplots()
line, = ax.plot([], [], lw=2)
ax.set_xlim(0, 0.05)  # zoom em 50 ms
ax.set_ylim(-3, 3)
ax.set_title("Animação: Onda composta (D4, G4, B4)")

def update(frame):
    i = frame * 500  # 500 amostras por frame
    x = t[i:i+500]
    y = wave_sum[i:i+500]
    line.set_data(x, y)
    return line,

ani = FuncAnimation(fig, update, frames=range(0, int(len(t) / 2000)), blit=True)
ani.save("acorde_sol_maior.gif", writer='pillow', fps=24)
plt.close()

# --- GRÁFICO GRANDE DO ACORDE COMPLETO ---
plt.figure(figsize=(12, 4))
plt.plot(t[:2000], wave_sum[:2000])
plt.title("Acorde de Sol Maior (G4, B4, D4)")
plt.xlabel("Tempo (s)")
plt.ylabel("Amplitude")
plt.ylim(-4, 4)
plt.grid(True)
plt.tight_layout()
plt.savefig("acorde_sol_maior.png", dpi=300)
plt.show()
