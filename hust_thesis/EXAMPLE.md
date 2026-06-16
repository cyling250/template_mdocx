# 绪论

## 研究背景与意义

华中科技大学是教育部直属重点**综合性大学**，是国家*双一流*建设高校。行内公式 $E = mc^2$ 是著名的质能方程。

如 @ref{fig:arch} 所示，本文提出的模型架构具有创新性。相关工作详见 @cite{resnet2016,transformer2017}。

### 国内外研究现状

基于CNN的方法 @cite{resnet2016} 在图像识别任务中表现出色 @cite{zhang2023}。

@formula[eq:accuracy]

$$
\text{Accuracy} = \frac{TP + TN}{TP + TN + FP + FN}
$$

基于Transformer的方法近年来取得了**显著成果**。

@formula[eq:attention]

$$
\text{Attention}(Q, K, V) = \text{softmax}\left(\frac{QK^T}{\sqrt{d_k}}\right)V
$$

### 本文贡献

本文的主要贡献包括以下三个方面。

@figure[fig:arch]{本文提出的端到端模型架构}
![架构图](E:\mcp-server\template_mdocx\docs\image.png)

## 相关工作

### 图像分类方法

传统的图像分类方法依赖于手工设计的特征。近年来，深度学习方法的引入极大提升了性能 @cite{resnet2016,transformer2017,vit2021}。

@formula[eq:loss]

$$
\mathcal{L} = \mathcal{L}_{\text{cls}} + \lambda \mathcal{L}_{\text{reg}}
$$

不同方法在ImageNet数据集上的性能对比如 @ref{tbl:comparison} 所示。

@table[tbl:comparison]{不同方法在ImageNet上的性能对比}
| 方法 | Top-1准确率 | Top-5准确率 | 参数量 |
| --- | --- | --- | --- |
| ResNet-50 | 76.1% | 92.9% | 25.6M |
| ResNet-101 | 77.4% | 93.5% | 44.5M |
| ViT-B/16 | 77.9% | 93.7% | 86.6M |
| 本文方法 | 78.5% | 94.1% | 32.1M |

#### 消融实验

为了验证各模块的有效性，我们进行了消融实验 @cite{resnet2016,transformer2017,zhang2023,pytorch2019}，结果如 @ref{tbl:ablation} 所示。

@table[tbl:ablation]{消融实验结果}
| 配置 | 准确率 | 参数量 |
| --- | --- | --- |
| Baseline | 72.3% | 20.1M |
| +模块A | 75.6% | 25.3M |
| +模块A+B | 77.2% | 28.7M |
| +模块A+B+C | 78.5% | 32.1M |

## 结论与展望

本文提出了一种新颖的深度学习模型，在多个基准数据集上取得了**优异的性能**。优化目标函数如 @ref{eq:opt} 所示。

@formula[eq:opt]

$$
\min_{\theta} \frac{1}{N} \sum_{i=1}^{N} \mathcal{L}(f(x_i; \theta), y_i) + \lambda \Omega(\theta)
$$

### 未来工作

未来的工作将集中在以下几个方面：探索更高效的网络架构，将方法扩展到更多的应用场景，研究模型的可解释性。

实验采用以下配置：学习率为 $1 \times 10^{-4}$，批次大小为 $64$。不同输入分辨率下的性能表现如 @ref{tbl:resolution} 所示。

@table[tbl:resolution]{不同输入分辨率下的性能}
| 分辨率 | 准确率 | 推理时间 |
| --- | --- | --- |
| 224×224 | 78.5% | 12ms |
| 384×384 | 80.2% | 28ms |
| 512×512 | 81.0% | 48ms |

@bibliography{refs.bib}
