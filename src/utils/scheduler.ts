export interface Invigilator {
  id: string;
  name: string;
  experience?: number;
}

export interface AssignmentRules {
  minimizeConsecutive: boolean;
  fixedPairs: boolean;
  balanceExperience: boolean;
  avoidPairs?: [string, string][];
}

export const compareVietnameseName = (nameA: string, nameB: string) => {
  const getTen = (name: string) => {
    const parts = name.trim().split(/\s+/);
    return parts[parts.length - 1];
  };
  const tenA = getTen(nameA);
  const tenB = getTen(nameB);
  const cmp = tenA.localeCompare(tenB, 'vi');
  if (cmp !== 0) return cmp;
  return nameA.localeCompare(nameB, 'vi');
};

export interface Assignment {
  shift: number;
  room: number;
  invigilator1: Invigilator;
  invigilator2?: Invigilator;
}

export interface AssignmentResult {
  schedule?: Assignment[][];
  stats?: {
    invigilator: Invigilator;
    count: number;
  }[];
  error?: string;
}

export function generateSchedule(
  invigilators: Invigilator[],
  numShifts: number,
  numRooms: number,
  invigilatorsPerRoom: number = 2,
  rules: AssignmentRules = { minimizeConsecutive: false, fixedPairs: false, balanceExperience: false }
): AssignmentResult {
  const N = invigilators.length;
  const S = numShifts;
  const R = numRooms;

  if (N < invigilatorsPerRoom * R) {
    return { error: `Không đủ giám thị. Cần ít nhất ${invigilatorsPerRoom * R} giám thị cho ${R} phòng thi.` };
  }

  const totalSlots = invigilatorsPerRoom * R * S;
  const idealMax = Math.ceil(totalSlots / N);
  const idealMin = Math.floor(totalSlots / N);

  const avoidMap = new Map<string, Set<string>>();
  if (rules.avoidPairs) {
    rules.avoidPairs.forEach(([id1, id2]) => {
      if (!avoidMap.has(id1)) avoidMap.set(id1, new Set());
      if (!avoidMap.has(id2)) avoidMap.set(id2, new Set());
      avoidMap.get(id1)!.add(id2);
      avoidMap.get(id2)!.add(id1);
    });
  }

  const maxAttempts = 500;
  let bestSchedule: Assignment[][] | null = null;
  let bestStats: any = null;
  let bestScore = Infinity;

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const schedule: Assignment[][] = [];
    const shiftsCount = new Array(N).fill(0);
    const roomsVisited = Array.from({ length: N }, () => new Set<number>());
    const pairsWorked = Array.from({ length: N }, () => new Set<number>());
    const shiftsWorked = Array.from({ length: N }, () => new Set<number>());
    let attemptScore = 0;
    let success = true;

    for (let s = 0; s < S; s++) {
      const shiftSchedule: Assignment[] = [];
      const workingInShift = new Set<number>();

      for (let r = 0; r < R; r++) {
        const roomAssigned: number[] = [];

        for (let k = 0; k < invigilatorsPerRoom; k++) {
          let bestInv = -1;
          let bestInvScore = Infinity;

          // Randomize candidates to explore different assignments
          const candidates = Array.from({length: N}, (_, i) => i)
            .filter(i => !workingInShift.has(i))
            .sort(() => Math.random() - 0.5);

          const hasExperiencedAvailable = candidates.some(c => (invigilators[c].experience || 0) > 0);
          const roomHasExperienced = roomAssigned.some(a => (invigilators[a].experience || 0) > 0);

          for (const i of candidates) {
            let score = shiftsCount[i] * 10; // Base score to balance shifts

            // Heavy penalty for exceeding ideal max shifts
            if (shiftsCount[i] >= idealMax) {
              score += 10000;
            }

            // Penalty for visiting the same room
            if (roomsVisited[i].has(r)) {
              score += 1000;
            }

            // Penalty for consecutive shifts
            if (rules.minimizeConsecutive) {
              let consecutiveCount = 0;
              for (let prevS = s - 1; prevS >= 0; prevS--) {
                if (shiftsWorked[i].has(prevS)) {
                  consecutiveCount++;
                } else {
                  break;
                }
              }
              
              if (consecutiveCount === 1) {
                score += 3000; // Penalty for working 2 consecutive shifts
              } else if (consecutiveCount === 2) {
                score += 8000; // Heavy penalty for working 3 consecutive shifts
              } else if (consecutiveCount >= 3) {
                score += 20000; // Extreme penalty for 4+ consecutive shifts
              }
            }

            // Penalty/Reward for working with the same partner
            let pairViolation = false;
            let pairMatch = false;
            for (const assigned of roomAssigned) {
              if (pairsWorked[i].has(assigned)) {
                pairViolation = true;
                pairMatch = true;
                break;
              }
            }
            
            if (rules.fixedPairs) {
              if (pairMatch) {
                score -= 2000; // Reward working with the same partner
              } else if (roomAssigned.length > 0 && pairsWorked[i].size > 0) {
                // If there's already someone in the room, and this person hasn't worked with them,
                // but this person has worked with someone else before, penalize breaking pairs.
                score += 1500; 
              }
            } else {
              if (pairViolation) {
                score += 1000;
              }
            }
            
            // Priority by experience
            if (rules.balanceExperience) {
              const currentExp = invigilators[i].experience || 0;
              
              if (!roomHasExperienced) {
                // If the room doesn't have an experienced invigilator yet
                if (currentExp === 0 && hasExperiencedAvailable) {
                  score += 50000; // Extreme penalty: do not assign inexperienced if experienced is available
                } else if (currentExp > 0) {
                  score -= 2000; // Reward picking experienced
                }
              } else {
                // Room already has an experienced invigilator
                if (currentExp === 0) {
                  score -= 2000; // Reward picking inexperienced to balance
                } else {
                  score += 2000; // Penalty for putting multiple experienced in same room if we could balance
                }
              }
            }

            // Penalty for avoid pairs
            if (rules.avoidPairs && roomAssigned.length > 0) {
              const currentInvId = invigilators[i].id;
              let avoidViolation = false;
              for (const assigned of roomAssigned) {
                const assignedId = invigilators[assigned].id;
                if (avoidMap.get(currentInvId)?.has(assignedId)) {
                  avoidViolation = true;
                  break;
                }
              }
              if (avoidViolation) {
                score += 20000; // Very heavy penalty to strictly avoid
              }
            }

            if (score < bestInvScore) {
              bestInvScore = score;
              bestInv = i;
            }
          }

          if (bestInv === -1) {
            success = false;
            break;
          }

          roomAssigned.push(bestInv);
          workingInShift.add(bestInv);
          shiftsCount[bestInv]++;
          roomsVisited[bestInv].add(r);
          shiftsWorked[bestInv].add(s);
          attemptScore += bestInvScore;
        }

        if (!success) break;

        for (let i = 0; i < roomAssigned.length; i++) {
          for (let j = i + 1; j < roomAssigned.length; j++) {
            pairsWorked[roomAssigned[i]].add(roomAssigned[j]);
            pairsWorked[roomAssigned[j]].add(roomAssigned[i]);
          }
        }

        shiftSchedule.push({
          shift: s + 1,
          room: r + 1,
          invigilator1: invigilators[roomAssigned[0]],
          invigilator2: roomAssigned.length > 1 ? invigilators[roomAssigned[1]] : undefined
        });
      }
      if (!success) break;
      schedule.push(shiftSchedule);
    }

    if (success) {
      // Add penalty if minimum shifts are not met (to ensure balance)
      let maxShifts = 0;
      let minShifts = Infinity;
      for (let i = 0; i < N; i++) {
        maxShifts = Math.max(maxShifts, shiftsCount[i]);
        minShifts = Math.min(minShifts, shiftsCount[i]);
        if (shiftsCount[i] < idealMin) {
          attemptScore += 5000;
        }
      }
      
      const balancePenalty = (maxShifts - minShifts) * 100;
      attemptScore += balancePenalty;

      if (attemptScore < bestScore) {
        bestScore = attemptScore;
        bestSchedule = schedule;
        bestStats = invigilators.map((inv, idx) => ({
          invigilator: inv,
          count: shiftsCount[idx]
        })).sort((a, b) => {
          if (b.count !== a.count) return b.count - a.count;
          return compareVietnameseName(a.invigilator.name, b.invigilator.name);
        });
      }

      // If perfect schedule found (no major penalties and perfectly balanced)
      if (bestScore < 1000 && (maxShifts - minShifts) <= 1) {
        break;
      }
    }
  }

  if (bestSchedule) {
    return { schedule: bestSchedule, stats: bestStats };
  }

  return { error: "Không thể tìm được lịch phân công phù hợp. Vui lòng thử lại hoặc điều chỉnh số lượng giám thị/phòng thi." };
}
