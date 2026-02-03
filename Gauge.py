import matplotlib.pyplot as plt
import numpy as np

def create_gauge(value, title='Gauge Meter', min_val=0, max_val=100, segments=None):
    """
    Create a gauge meter.

    Parameters:
    - value: The current value to display.
    - title: The title of the gauge.
    - min_val: The minimum value of the gauge.
    - max_val: The maximum value of the gauge.
    - segments: List of tuples [(range_start, range_end, color)] for colored segments.
    """
    # Create a figure and axis
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw={'projection': 'polar'})
    
    # Hide the polar grid and labels
    ax.set_theta_offset(np.pi / 2)  # Start at 12 o'clock
    ax.set_theta_direction(-1)     # Draw clockwise
    ax.set_yticklabels([])         # Hide radial labels
    ax.set_xticklabels([])         # Hide angular labels
    ax.spines['polar'].set_visible(False)
    
    # Define the gauge segments
    if not segments:
        segments = [
            (0, 25, 'red'),
            (25, 50, 'orange'),
            (50, 75, 'yellow'),
            (75, 100, 'green')
        ]
    
    # Draw the segments
    for start, end, color in segments:
        start_angle = np.deg2rad(180 * start / max_val)
        end_angle = np.deg2rad(180 * end / max_val)
        ax.barh(0, [1], left=[start_angle], width=[end_angle - start_angle],
                color=color, edgecolor='black', height=0.5)
    
    # Draw the needle
    needle_angle = np.deg2rad(180 * value / max_val)
    ax.plot([0, needle_angle], [0, 0.5], color='black', linewidth=2)
    
    # Add labels
    for percent in [0, 25, 50, 75, 100]:
        angle = np.deg2rad(180 * percent / max_val)
        ax.text(angle, 0.65, f'{percent}%', horizontalalignment='center', verticalalignment='center')
    
    # Add title
    ax.text(0, 1.2, title, horizontalalignment='center', fontsize=14, fontweight='bold')
    
    # Show the gauge
    plt.show()

# Call the function
create_gauge(50, title='Progress')