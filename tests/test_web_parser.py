import sys
import time

def loading_bar():
    total_steps = 50
    for step in range(total_steps + 1):
        sys.stdout.write('\r')
        sys.stdout.write(f"[{'=' * step}{' ' * (total_steps - step)}] {step * 2}%")
        sys.stdout.flush()
        time.sleep(0.1)  # Adjust sleep time to control the speed of the loading bar

    print("\nLoading complete!")

if __name__ == "__main__":
    loading_bar()
