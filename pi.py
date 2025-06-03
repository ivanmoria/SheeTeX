import matplotlib.pyplot as plt

# Dados para o gráfico
dinamicas = ['F', 'mp', 'mf', 'p']  # As dinâmicas em notação musical
intensidade_dB = [95, 55, 75, 35]  # Intensidades em decibéis (dB)

# Definir estilo moderno para o gráfico
plt.style.use('fivethirtyeight')  # Usando estilo fivethirtyeight

# Criando o gráfico
plt.figure(figsize=(8, 5))

# Plotando o gráfico
plt.plot(dinamicas, intensidade_dB, marker='o', color='#000000', linestyle='-', linewidth=3, markersize=10, markerfacecolor='white', markeredgewidth=2)

plt.xlabel('Dinâmica', fontsize=14, color='#555555', style='italic', weight='bold')
plt.ylabel('Intensidade (dB)', fontsize=14, color='#555555')

# Ajustando os limites do eixo Y
plt.ylim(0, 110)

# Ajustando a fonte das dinâmicas no eixo X para uma fonte musical
plt.xticks(fontsize=14, color='#444444', style='italic', fontweight='bold', fontname='DejaVu Sans')

# Tornar os ticks dos eixos mais visíveis
plt.yticks(fontsize=12, color='#444444')

# Removendo o grid do fundo (já removido pelo estilo)
plt.grid(False)

# Garantir que o fundo da figura seja transparente
plt.gcf().patch.set_facecolor('none')
plt.gca().set_facecolor('none')

# Ajustando o layout automaticamente
plt.tight_layout()

# Salvando em PNG com fundo transparente
plt.savefig('grafico_dinamicas_transparente.png', format='png', dpi=300, transparent=True)

# Exibindo o gráfico
plt.show()
