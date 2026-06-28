# M12 · 에이전트로 영상 만들기 (Remotion)

## 실습 목표

에이전트로 Remotion 프로젝트를 만들고 5초 소개 타이틀 영상을 MP4로 렌더링합니다.

## 실습

Node.js 버전을 확인한 뒤 빈 프로젝트 폴더에서 요청합니다.

```text
Remotion 프로젝트를 새로 만들고 1920×1080, 30fps, 5초짜리 소개 타이틀 영상을 제작해줘.
이름과 한 줄 소개가 순서대로 등장하고, spring 또는 interpolate를 사용해 자연스럽게 움직이게 해줘.
Remotion Studio 미리보기와 MP4 렌더 명령도 알려줘.
```

프로젝트가 생성되면 에이전트가 안내한 위치에서 실행합니다.

```powershell
npx remotion studio
npx remotion render
```

텍스트, 색상, 등장 시점을 한 번 수정한 뒤 다시 렌더링합니다.

## 완료 확인

- [ ] Remotion 프로젝트를 생성했다.
- [ ] Studio에서 5초 영상을 미리봤다.
- [ ] 애니메이션을 한 번 수정했다.
- [ ] MP4 렌더링을 완료했다.
