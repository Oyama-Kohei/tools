CREATE TABLE t_users
(
    user_id character varying(36) COLLATE pg_catalog."default" NOT NULL,
    password character varying(36) COLLATE pg_catalog."default" NOT NULL,
    created_at timestamp with time zone NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_updated_at timestamp with time zone NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT t_users_pkey PRIMARY KEY (user_id, password)
)

TABLESPACE pg_default;

ALTER TABLE t_users OWNER to postgres;

COMMENT ON TABLE t_users IS 'ユーザ管理';

COMMENT ON COLUMN t_users.user_id IS 'ユーザID';

COMMENT ON COLUMN t_users.password IS 'ユーザパスワード';

COMMENT ON COLUMN t_users.created_at IS 'データ作成日';

COMMENT ON COLUMN t_users.last_updated_at IS 'データ更新日';