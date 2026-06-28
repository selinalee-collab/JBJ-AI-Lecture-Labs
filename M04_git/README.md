# M04 · Git 기초

## 실습 목표

강의 저장소를 Fork하고 Clone, Commit, Push, Pull을 한 바퀴 수행합니다.

## 실습 순서

1. 이 저장소를 본인 GitHub 계정으로 Fork합니다.
2. 본인의 Fork URL로 저장소를 Clone합니다.

```powershell
git clone https://github.com/<내아이디>/JBJ-AI-Lecture-Labs.git
cd JBJ-AI-Lecture-Labs
git status
```

3. 루트에 `hello_본인이름.txt`를 만들고 간단한 자기소개를 적습니다.
4. 변경을 커밋하고 Fork에 올립니다.

```powershell
git add hello_본인이름.txt
git commit -m "실습: 내 소개 파일 추가"
git push
git pull
```

5. GitHub 웹에서 파일과 커밋을 확인합니다.
6. 선택 실습으로 [`git_training/main.txt`](./git_training/main.txt)를 수정하고 `git diff`로 변경 전후를 확인합니다.

## 완료 확인

- [ ] Fork와 Clone을 완료했다.
- [ ] `git status`와 `git diff`를 확인했다.
- [ ] 소개 파일을 Commit하고 Push했다.
- [ ] GitHub 웹에서 파일을 확인하고 Pull했다.
