-- ════════════════════════════════════════════════════════════════
-- 포레온 홈페이지 보안 강화 (Supabase → SQL Editor 에 통째로 붙여넣고 Run)
-- 반드시 먼저: ① 회원가입 끄기 ② Authentication 사용자 만들기 ③ 새 로그인 화면 배포
-- 실행하면: 로그인한 직원만 표를 읽고 쓸 수 있고, 키만 가진 외부인은 아무것도 못 봅니다.
-- ════════════════════════════════════════════════════════════════
do $$
declare r record;
begin
  -- 1) 기존 정책(누구나 허용 등) 모두 정리
  for r in select policyname, tablename from pg_policies where schemaname = 'public' loop
    execute format('drop policy %I on public.%I', r.policyname, r.tablename);
  end loop;
  -- 2) 모든 표: 보안(RLS) 켜고, 로그인한 직원만 허용
  for r in select tablename from pg_tables where schemaname = 'public' loop
    execute format('alter table public.%I enable row level security', r.tablename);
    execute format('create policy staff_only on public.%I for all to authenticated using (true) with check (true)', r.tablename);
  end loop;
end $$;

-- 3) 파일 저장소(shared)도 비공개 + 로그인한 직원만
update storage.buckets set public = false where id = 'shared';
drop policy if exists staff_only_shared on storage.objects;
create policy staff_only_shared on storage.objects for all to authenticated
  using (bucket_id = 'shared') with check (bucket_id = 'shared');

-- 확인: 아래 결과의 rowsecurity 가 모두 true 이면 완료
select tablename, rowsecurity from pg_tables where schemaname = 'public' order by tablename;
