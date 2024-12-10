import numpy as np
import random

# 定义环境：迷宫的大小
n_rows, n_cols = 4, 4
goal = (3, 3)  # 目标位置
obstacles = [(1, 1), (2, 1)]  # 障碍物位置

# 定义动作：上、下、左、右
actions = ['up', 'down', 'left', 'right']

# 初始化 Q 表：每个状态-动作对的 Q 值
Q = np.zeros((n_rows, n_cols, len(actions)))

# 学习率和折扣因子
alpha = 0.1
gamma = 0.9
epsilon = 0.1  # epsilon-greedy 策略中的 epsilon



# epsilon-greedy 策略
def choose_action(state):
    if random.uniform(0, 1) < epsilon:
        return random.choice(actions)  # 随机选择动作
    else:
        r, c = state
        action_values = Q[r, c, :]
        max_action = np.argmax(action_values)
        return actions[max_action]  # 选择 Q 值最大的动作

# 训练过程
def train(episodes=1000):
    for _ in range(episodes):
        state = (0, 0)  # 从起始位置开始
        action = choose_action(state)
        next_state=()
        reward=0
        while state != goal:
            # 执行动作，得到新状态和奖励
           
            if action == 'up': next_state = (state[0] - 1, state[1])
            elif action == 'down': next_state = (state[0] + 1, state[1])
            elif action == 'left': next_state = (state[0], state[1] - 1)
            else: next_state = (state[0], state[1] + 1)
            
            if next_state[0] < 0 or next_state[0] >= n_rows or next_state[1] < 0 or next_state[1] >= n_cols:
                next_state = state  # 如果越界，保持当前状态s
            if next_state==(1,1)  or next_state==(2,1):
                reward=-5
                next_state = state
            # 奖励：如果达到目标则奖励 +1，否则为 -1
            else: reward = 0
            if next_state == goal:
                reward = 100

            # 选择下一个动作
            next_action = choose_action(next_state)

            # 更新 Q 值
            r, c = state
            next_r, next_c = next_state
            action_idx = actions.index(action)
            next_action_idx = actions.index(next_action)
            Q[r, c, action_idx] += alpha * (reward + gamma * Q[next_r, next_c, next_action_idx] - Q[r, c, action_idx])

            # 更新状态和动作
            state, action = next_state, next_action
            

# 训练
train(10000)

# 输出学习到的 Q 表
print(Q)
