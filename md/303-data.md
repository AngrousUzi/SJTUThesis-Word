# 数据

: 推介演讲软信息度量指标

|            维度            |                                 标记                                 | 计算方法                                                                                                                                          |
| :------------------------: | :-------------------------------------------------------------------: | ------------------------------------------------------------------------------------------------------------------------------------------------- |
|   **视觉（情绪）**   |                   $Emo_{happy}$, $Emo_{sad}$ 等                   | 人脸帧上的平均预测概率：$\frac{1}{N_{face}} \sum_{i=1}^{N_{face}} P(\text{emo}_i)$。通过 HSEmotion EfficientNet 提取。                          |
|                            |                               $PosER$                               | ${emo\_dominant}=argmax(Emo_{facial})$。序列中*主导*情绪为*愉快*或*惊讶*的帧比例：$\frac{N_{pos\_dominant}}{N_{face}}$。                |
|                            |                               $NegER$                               | 序列中*主导*情绪为*愤怒、恐惧、厌恶、悲伤、蔑视*的帧比例：$\frac{N_{neg\_dominant}}{N_{face}}$。                                            |
|                            |                               $NeuER$                               | 序列中*主导*情绪为*中性*的帧比例：$\frac{N_{neu\_dominant}}{N_{face}}$。                                                                    |
|   **视觉（注视）**   |                              $GazeCam$                              | $\frac{1}{N_{face}} \sum \mathbb{I}(\|yaw\| < 5^\circ \land \|gaze_x\| < 0.15 \land \|gaze_y\| < 0.15)$。要求正面姿态和居中虹膜的严格联合度量。 |
|                            |                            $HeadYawStd$                            | 头部偏航角（水平旋转）的标准差：$\sigma(\theta_{yaw})$。                                                                                        |
|                            |                           $HeadPitchStd$                           | 头部俯仰角（垂直倾斜）的标准差：$\sigma(\theta_{pitch})$。                                                                                      |
|  **语音（流利度）**  |                               $ArtR$                               | $\frac{\text{总字符数}}{\sum (t_{word\_end} - t_{word\_start})}$。字符数除以实际单词发音时间，不包括句内停顿。                                  |
|                            |                              $PauseR$                              | $\frac{T_{total} - \sum (t_{seg\_end} - t_{seg\_start})}{T_{total}}$。静音/非说话音频段占总演示时长的比例。                                     |
| **语音（动态变化）** |                              $f0_{cv}$                              | $\frac{\sigma(F0)}{\mu(F0)}$。通过 pYIN 仅在有声帧（$\mathbb{I}_{voiced}=1$）提取的基频。                                                     |
|                            |                            $f0_{range}$                            | 基频的四分位距：$P_{75}(F0) - P_{25}(F0)$。相比于最大最小值边界，对声学异常值更具鲁棒性。                                                       |
|                            |                            $f0_{slope}$                            | 线性 OLS 回归的系数$\beta_1$：$F0_t = \beta_0 + \beta_1 t +\varepsilon_t$，在有效的有声帧上计算。                                             |
|                            |                             $RMS_{cv}$                             | $\frac{\sigma(RMS)}{\mu(RMS)}$，在短时窗口内计算的均方根（RMS）能量的变异系数。                                                                 |
|                            |                            $RMS_{range}$                            | $P_{95}(RMS) - P_{05}(RMS)$。计算出的音频能量幅度的 90% 中心概率跨度。                                                                          |
|   **文本（情感）**   | $PosAnnR_{p}$ / $NegAnnR_{p}$ / $PosSclR_{p}$ / $NegSclR_{p}$ | $\frac{N_{\text{正面/负面词数}}}{N_{\text{总词数}}}$，基于在相应财务/社交词典中的精确匹配。                                                     |

注：所有连续控制变量在进入回归模型前均已在 1% 和 99% 分位数处进行缩尾处理（winsorized）。

![yearly average CAR](image/303-data/CAR.png)

*Note: Year-by-year average rival CAR across the three 30-minute event windows (before, during, after). The dashed vertical line marks the 2016 subscription-system reform. Pre-reform years exhibit a uniformly negative placebo CAR consistent with the mechanical capital-substitution channel. While post-reform years display a noticeably wider cross-year spread in the during- and after-roadshow windows, motivating the subsequent conditional analysis.*

To investigate whether the persuasiveness of an IPO roadshow pitch exerts a direct spillover effect on listed rivals, I estimate a baseline OLS regression model:

$$
Y_i = \alpha + \sum_{k=1}^{K} \beta_k \cdot PC_k^i + \gamma \cdot \mathbf{Z}_i + \delta_t + \varepsilon_i
$$

where $CAR_{j,i,[T_1,T_2]}$ is the cumulative abnormal return of rival firm $j$ during event window $[T_1, T_2]$ around IPO $i$'s roadshow timing; $PC_k^i$ are the multimodal Pitch Factor scores extracted via PCA from the roadshow's visual, vocal, and verbal dimensions, entered iteratively across columns (1)–(3); $\mathbf{Z}^{rival}_{j}$ is a vector of rival-firm characteristics (log market cap, book-to-market, ROA, leverage, established age and listing age); $\mathbf{Z}^{IPO}_{i}$ is a vector of IPO-level characteristics (log offering size, diluted P/E ratio, shares issued, and issue price and roadshow duration); and $sim\_mda_{j,i}$ is the pairwise TF-IDF cosine similarity between the MD\&A texts of the IPO firm and the rival. The sample is restricted to the top-3 most textually similar rivals per IPO. The fixed effects $\delta_t$ comprise listing year, CSRC 3-digit industry, listing board (STAR / ChiNext / SH Main / SZ Main), and roadshow platform. Standard errors are clustered at the IPO level to account for within-event correlation across rival observations. All continuous variables are winsorized at the 1st and 99th percentiles except PCA components which has been winsorized. The regression is run separately for three event sub-windows—"Before Start" (placebo), "After Start", and "After End"—and for each of the AM (09:00), PM (14:00), and pooled sub-samples.
